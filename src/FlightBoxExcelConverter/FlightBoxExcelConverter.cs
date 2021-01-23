using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CsvHelper;
using FlightBoxExcelConverter.Enums;
using FlightBoxExcelConverter.Exporter;
using FlightBoxExcelConverter.Objects;
using FlightBoxExcelConverter.Properties;

namespace FlightBoxExcelConverter
{
    public class FlightBoxExcelConverter
    {
        private readonly bool _ignoreDateRange;
        public event EventHandler<LogEventArgs> LogEventRaised;
        public event EventHandler ExportFinished;
        private List<string> _logEntries = new List<string>();

        public string ImportFileName { get; set; }

        public string ExportFolderName { get; set; }

        private DateTime CreationTimeStamp { get; set; }

        public bool HasExportError { get; set; }

        public string ExportErrorMessage { get; set; }

        private readonly DataManager _dataManager;

        public FlightBoxExcelConverter(string importFileName, string exportFolderName, bool ignoreDateRange)
        {
            _ignoreDateRange = ignoreDateRange;
            ImportFileName = importFileName;
            ExportFolderName = exportFolderName;
            _dataManager = new DataManager();
        }

        public DateTime GetLastWriteDateTimeOfImportFileName()
        {
            if (ImportFileName.ToLower().EndsWith(".csv") == false)
            {
                var directory = new DirectoryInfo(ImportFileName);
                var lastFile = directory.GetFiles().OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                if (lastFile == null)
                {
                    OnLogEventRaised("Konnte keine aktuelle Datei zum Importieren finden!");
                    throw new ApplicationException($"Konnte keine aktuelle Datei zum Importieren in Verzeichnis {ImportFileName} finden!");
                }

                ImportFileName = lastFile.FullName;
            }

            if (File.Exists(ImportFileName))
            {
                var fi = new FileInfo(ImportFileName);
                return fi.LastWriteTime;
            }

            //no file to import found or empty file string selected
            throw new ApplicationException("Konnte keine Datei zum Importieren finden oder es wurde keine Datei zum Importieren ausgewählt!");
        }
        
        public void Convert()
        {
            try
            {
                ExportErrorMessage = string.Empty;
                HasExportError = false;

                //loading proffix address numbers from database
                var nrOfAddressesFromProffixRead = _dataManager.ReadProffixDatabase();

                if (ImportFileName.ToLower().EndsWith(".csv") == false)
                {
                    var directory = new DirectoryInfo(ImportFileName);
                    var lastFile = directory.GetFiles().OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                    if (lastFile == null)
                    {
                        OnLogEventRaised("Konnte keine aktuelle Datei zum Importieren finden!");
                        throw new ApplicationException($"Konnte keine aktuelle Datei zum Importieren in Verzeichnis {ImportFileName} finden!");
                    }

                    ImportFileName = lastFile.FullName;
                    OnLogEventRaised($"Verwende Datei: {ImportFileName}");
                }

                _logEntries = new List<string>();
                var flightBoxDataList = ReadFile();
                var proffixDataList = new List<ProffixData>();
                CreationTimeStamp = DateTime.Now;

                OnLogEventRaised("Bereinige Basis-Daten...");

                if (Settings.Default.ReadProffixDbData)
                {
                    OnLogEventRaised($"Mitgliedernummern werden anhand der Proffix-Datenbank überprüft. {nrOfAddressesFromProffixRead} Adressen aus Proffix geladen.");
                }
                else
                {
                    OnLogEventRaised("Proffix-Datenbank wird für die Mitgliedernummer-Prüfung nicht verwendet!");
                }

                foreach (var flightBoxData in flightBoxDataList)
                {
                    var proffixData = new ProffixData(flightBoxData);

                    //check movement date within the valid import range
                    //we won't import old movement data into proffix
                    var minMovementDate = new DateTime(DateTime.Now.AddMonths(-1).Year, DateTime.Now.AddMonths(-1).Month, 1);

                    if (flightBoxData.MovementDateTime < minMovementDate)
                    {
                        // found movement older than the previous months
                        flightBoxData.IsOlderMovementDate = true;
                        OnLogEventRaised(
                            $"Alte Flugbewegung gefunden vom {flightBoxData.MovementDateTime.ToShortDateString()} (Zeile: {flightBoxData.LineNumber}).");

                        if (_ignoreDateRange == false)
                        {
                            ExportErrorMessage = "Alte Flugbewegung wurde gefunden und Warnung bei alten Daten darf nicht ignoriert werden. Verarbeitung wird abgebrochen.";
                            HasExportError = true;
                            WriteLogFile();
                            ExportFinished?.Invoke(this, EventArgs.Empty);
                            return;
                        }
                    }

                    // try to find MemberNumber based on name if no MemberNumber is set
                    if (string.IsNullOrWhiteSpace(flightBoxData.MemberNumber) || flightBoxData.MemberNumber == "000000")
                    {
                        if (_dataManager.FindLastnameAndSetMemberNumber(proffixData))
                        {
                            OnLogEventRaised(
                                $"MemberNumber {proffixData.MemberNumber} für {flightBoxData.Lastname} mit {flightBoxData.Immatriculation} gesetzt (Zeile: {flightBoxData.LineNumber}).");
                        }
                    }

                    // set MemberNumber based on immatriculation
                    if (_dataManager.FindImmatriculationAndMapMemberNumber(proffixData))
                    {
                        OnLogEventRaised($"Setze spezielle Mitgliedernummer für {proffixData.FlightBoxData.Immatriculation} (Zeile: {flightBoxData.LineNumber}): Alte Mitgliedernummer {proffixData.FlightBoxData.MemberNumber}, neue Mitgliedernummer {proffixData.MemberNumber}");
                    }

                    if (Settings.Default.ReadProffixDbData && _dataManager.FindMemberNumberInProffix(proffixData) == false)
                    {
                        OnLogEventRaised($"Mitgliedernummer {proffixData.MemberNumber} für {proffixData.FlightBoxData.Lastname} mit {proffixData.FlightBoxData.Immatriculation} in Proffix-Datenbank nicht gefunden (Zeile: {flightBoxData.LineNumber})");
                    }

                    proffixDataList.Add(proffixData);
                }

                Thread.Sleep(50);

                WriteBaseFile(proffixDataList);

                Thread.Sleep(50);

                foreach (var proffixData in proffixDataList)
                {
                    if (_dataManager.IsNoLdgTaxMember(proffixData))
                    {
                        proffixData.FlightBoxData.IgnoreLandingTax = true;
                    }

                    if (proffixData.MemberNumber == "999605" &&
                        proffixData.FlightBoxData.TypeOfTraffic == (int) TypeOfTraffic.Instruction)
                    {
                        // Heli Sitterdorf is always private tax not training
                        proffixData.FlightBoxData.TypeOfTraffic = (int) TypeOfTraffic.Private;
                    }

                    if ((proffixData.MemberNumber == "999998" || proffixData.MemberNumber == "383909") &&
                        proffixData.FlightBoxData.TypeOfTraffic == (int)TypeOfTraffic.Instruction)
                    {
                        // Stoffel Aviation is external on instruction
                        proffixData.FlightBoxData.IsHomebased = false;
                    }

                    //TODO: Handle maintenance flights from Seiferle

                    // filtering for tow flights and departure movements are handled within the FlightBoxData class directly

                    if (string.IsNullOrWhiteSpace(proffixData.MemberNumber) || proffixData.MemberNumber == "000000")
                    {
                        proffixData.MemberNumberInProffixNotFound = true;

                        ExportErrorMessage +=
                            $"{Environment.NewLine}Fehlerhafte MemberNumber {proffixData.MemberNumber} für {proffixData.FlightBoxData.Lastname} mit {proffixData.FlightBoxData.Immatriculation} gefunden (Zeile: {proffixData.FlightBoxData.LineNumber}).";
                        OnLogEventRaised(
                            $"Fehlerhafte MemberNumber {proffixData.MemberNumber} für {proffixData.FlightBoxData.Lastname} mit {proffixData.FlightBoxData.Immatriculation} gefunden (Zeile: {proffixData.FlightBoxData.LineNumber}).");
                    }

                    CalculateLandingTax(proffixData);
                }

                if (HasExportError)
                {
                    WriteLogFile();
                    ExportFinished?.Invoke(this, EventArgs.Empty);
                    return;
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);
                var folder = Path.Combine(ExportFolderName, CreationTimeStamp.ToString("yyyy-MM-dd"));
                var exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_LdgTaxes_without_Remarks (to Import in Proffix).csv");
                var listToExport = proffixDataList.Where(x => x.FlightBoxData.IsDepartureMovement == false &&
                                                              x.FlightBoxData.IsMaintenanceFlight == false &&
                                                              x.FlightBoxData.IsTowFlight == false &&
                                                              x.FlightBoxData.IgnoreLandingTax == false &&
                                                              string.IsNullOrWhiteSpace(x.FlightBoxData.Remarks) &&
                                                              x.MemberNumberInProffixNotFound == false)
                                                              .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised($"Exportiere Proffix-Daten ohne Bemerkungen in Datei: {exportFilename}");
                    var exporter = new ProffixDataCsvExporter(exportFilename, listToExport);
                    exporter.Export();
                    OnLogEventRaised(
                        $"{exporter.NumberOfLinesExported} Proffix-Daten ohne Bemerkungen erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Proffix-Daten ohne Bemerkungen exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);

                exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_LdgTaxes_With_MemberNumberNotFound_Error (to correct and Import in Proffix).csv");
                listToExport = proffixDataList.Where(x => x.FlightBoxData.IsDepartureMovement == false &&
                                                              x.FlightBoxData.IsMaintenanceFlight == false &&
                                                              x.FlightBoxData.IsTowFlight == false &&
                                                              x.FlightBoxData.IgnoreLandingTax == false &&
                                                              x.MemberNumberInProffixNotFound)
                    .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised($"Exportiere Proffix-Daten mit fehlerhaften Mitgliedernummern in Datei: {exportFilename}");
                    var exporter = new ProffixDataCsvExporter(exportFilename, listToExport);
                    exporter.Export();
                    OnLogEventRaised(
                        $"{exporter.NumberOfLinesExported} Proffix-Daten mit fehlerhaften Mitgliedernummern erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Proffix-Daten mit fehlerhaften Mitgliedernummern exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);

                exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_LdgTaxes_with_Remarks (to check and import in Proffix).csv");
                listToExport = proffixDataList.Where(x => x.FlightBoxData.IsDepartureMovement == false &&
                                                          x.FlightBoxData.IsMaintenanceFlight == false &&
                                                          x.FlightBoxData.IsTowFlight == false &&
                                                          x.FlightBoxData.IgnoreLandingTax == false && 
                                                          string.IsNullOrWhiteSpace(x.FlightBoxData.Remarks) == false &&
                                                          x.MemberNumberInProffixNotFound == false)
                    .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised(
                        $"Exportiere Daten mit Bemerkungen zur Prüfung und Importieren in Proffix in Datei: {exportFilename}");
                    var exporter = new ProffixDataCsvExporter(exportFilename, listToExport);
                    exporter.Export();
                    OnLogEventRaised(
                        $"{exporter.NumberOfLinesExported} Daten mit Bemerkungen zur Prüfung und Importieren in Proffix erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Daten mit Bemerkungen zur Prüfung und Importieren in Proffix exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);

                exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_No_LdgTaxes_without_Remarks (not to import).csv");
                listToExport = proffixDataList.Where(x => (x.FlightBoxData.IsMaintenanceFlight ||
                                                           x.FlightBoxData.IsTowFlight ||
                                                           x.FlightBoxData.IgnoreLandingTax) &&
                                                          string.IsNullOrWhiteSpace(x.FlightBoxData.Remarks))
                    .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised($"Exportiere Nicht-Proffix-Daten ohne Bemerkungen in Datei: {exportFilename}");
                    var exporter = new ProffixDataCsvExporter(exportFilename, listToExport);
                    exporter.Export();
                    OnLogEventRaised(
                        $"{exporter.NumberOfLinesExported} Nicht-Proffix-Daten ohne Bemerkungen erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Nicht-Proffix-Daten ohne Bemerkungen exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);

                exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_No_LdgTaxes_with_Remarks (to check and NO import in Proffix).csv");
                listToExport = proffixDataList.Where(x => (x.FlightBoxData.IsMaintenanceFlight ||
                                                           x.FlightBoxData.IsTowFlight ||
                                                           x.FlightBoxData.IgnoreLandingTax) &&
                                                          string.IsNullOrWhiteSpace(x.FlightBoxData.Remarks) == false)
                    .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised(
                        $"Exportiere Daten mit Bemerkungen zur Prüfung und NICHT importieren in Proffix in Datei: {exportFilename}");
                    var exporter = new ProffixDataCsvExporter(exportFilename, listToExport);
                    exporter.Export();
                    OnLogEventRaised(
                        $"{exporter.NumberOfLinesExported} Daten mit Bemerkungen zur Prüfung und NICHT importieren in Proffix erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Daten mit Bemerkungen zur Prüfung und NICHT importieren in Proffix exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);

                exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_Heli Sitterdorf.csv");
                listToExport = proffixDataList.Where(x => x.MemberNumber == "999605" && x.FlightBoxData.IsDepartureMovement == false)
                    .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised($"Exportiere Daten für Heli Sitterdorf in Datei: {exportFilename}");
                    var reportExporter = new ReportExporter(exportFilename, listToExport);
                    reportExporter.Export();
                    OnLogEventRaised(
                        $"{reportExporter.NumberOfLinesExported} Daten für Heli Sitterdorf erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Daten für Heli Sitterdorf exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);

                exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_Skydive.csv");
                listToExport = proffixDataList.Where(x => x.MemberNumber == "703100" && x.FlightBoxData.IsDepartureMovement == false)
                    .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised($"Exportiere Daten für Skydive in Datei: {exportFilename}");
                    var reportExporter = new ReportExporter(exportFilename, listToExport);
                    reportExporter.Export();
                    OnLogEventRaised(
                        $"{reportExporter.NumberOfLinesExported} Daten für Skydive erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Daten für Skydive exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);
                OnLogEventRaised(string.Empty);

                exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_Swiss oldies.csv");
                listToExport = proffixDataList.Where(x => x.MemberNumber == "28" && x.FlightBoxData.IsDepartureMovement == false)
                    .ToList();

                if (listToExport.Any())
                {
                    OnLogEventRaised($"Exportiere Daten für Swiss oldies in Datei: {exportFilename}");
                    var reportExporter = new ReportExporter(exportFilename, listToExport);
                    reportExporter.Export();
                    OnLogEventRaised(
                        $"{reportExporter.NumberOfLinesExported} Daten für Swiss oldies erfolgreich exportiert.");
                }
                else
                {
                    OnLogEventRaised(
                        $"Keine Daten für Swiss oldies exportiert. Datei: {exportFilename} wurde nicht erzeugt!");
                }

                Thread.Sleep(50);

                WriteLogFile();

                Thread.Sleep(50);

                ExportFinished?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception e)
            {
                HasExportError = true;
                ExportErrorMessage = e.Message;
                OnLogEventRaised("Fehler beim Convertieren..." + Environment.NewLine + "Fehlermeldung: " + e.Message);
                ExportFinished?.Invoke(this, EventArgs.Empty);
            }
        }

        private void CalculateLandingTax(ProffixData proffixData)
        {
            // set Article number in Proffix data
            if (proffixData.FlightBoxData.TypeOfTraffic == (int)TypeOfTraffic.Instruction)
            {
                proffixData.ArticleNr = "1039"; //Landetaxen Speck Schulung
            }
            else
            {
                proffixData.ArticleNr = "1037"; //Landetaxen Speck (Charter)
            }

            // calculate quantity of landings in Proffix data
            if (proffixData.FlightBoxData.MovementType == "A") //Arrival
            {
                proffixData.ArticleQuantity = System.Convert.ToDecimal((proffixData.FlightBoxData.NrOfMovements + 1) / 2);
            }
            else if (proffixData.FlightBoxData.MovementType == "V") //circuits
            {
                proffixData.ArticleQuantity = System.Convert.ToDecimal(proffixData.FlightBoxData.NrOfMovements / 2);
            }
            else
            {
                proffixData.ArticleQuantity = 0;
            }

            proffixData.VfsArticleNumber = "1003";
            proffixData.VfsPrice = 1.0m;
            proffixData.VfsQuantity = proffixData.ArticleQuantity; //is same formula as for landing tax quantity calculation

            // calculate price for SchSpeck
            if (proffixData.FlightBoxData.IsHomebased && proffixData.FlightBoxData.TypeOfTraffic == (int)TypeOfTraffic.Instruction)
            {
                proffixData.SchHome = 0;
            }
            else
            {
                proffixData.SchHome = 0;
            }

            // calculate price for SchFremd
            if (proffixData.FlightBoxData.TypeOfTraffic == (int)TypeOfTraffic.Instruction
                && proffixData.FlightBoxData.IsHomebased == false)
            {
                if (proffixData.FlightBoxData.MaxTakeOffWeight < 1001)
                {
                    proffixData.SchExternal = 8.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1251)
                {
                    proffixData.SchExternal = 12.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1501)
                {
                    proffixData.SchExternal = 15.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 2001)
                {
                    proffixData.SchExternal = 20.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight > 2000)
                {
                    proffixData.SchExternal = 30.0m;
                }
                else
                {
                    proffixData.SchExternal = 0;
                }
            }
            else
            {
                proffixData.SchExternal = 0;
            }


            // calculate price for HB
            if (proffixData.FlightBoxData.IsHomebased
                && proffixData.FlightBoxData.TypeOfTraffic != (int)TypeOfTraffic.Instruction)
            {
                if (proffixData.FlightBoxData.MaxTakeOffWeight < 751)
                {
                    proffixData.LdgTaxHomebased = 12.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1001)
                {
                    proffixData.LdgTaxHomebased = 15.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1251)
                {
                    proffixData.LdgTaxHomebased = 17.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1501)
                {
                    proffixData.LdgTaxHomebased = 20.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 2001)
                {
                    proffixData.LdgTaxHomebased = 25.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight > 2000)
                {
                    proffixData.LdgTaxHomebased = 35.0m;
                }
                else
                {
                    proffixData.LdgTaxHomebased = 0;
                }
            }
            else
            {
                proffixData.LdgTaxHomebased = 0;
            }


            // calculate price for Fremd
            if (proffixData.FlightBoxData.IsHomebased == false
                && proffixData.FlightBoxData.TypeOfTraffic != (int)TypeOfTraffic.Instruction)
            {
                if (proffixData.FlightBoxData.MaxTakeOffWeight < 751)
                {
                    proffixData.LdgTaxExternal = 17.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1001)
                {
                    proffixData.LdgTaxExternal = 20.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1251)
                {
                    proffixData.LdgTaxExternal = 22.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1501)
                {
                    proffixData.LdgTaxExternal = 25.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 2001)
                {
                    proffixData.LdgTaxExternal = 30.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight > 2000)
                {
                    proffixData.LdgTaxExternal = 40.0m;
                }
                else
                {
                    proffixData.LdgTaxExternal = 0;
                }
            }
            else
            {
                proffixData.LdgTaxExternal = 0;
            }

            if (proffixData.FlightBoxData.Immatriculation == "YLLEV" ||
                proffixData.FlightBoxData.Immatriculation == "LYTED" ||
                proffixData.FlightBoxData.Immatriculation == "LYMHC")
            {
                //Antonov AN-2
                proffixData.LdgTaxExternal = 80.0m;
            }

            if (proffixData.FlightBoxData.Immatriculation == "HBZNP" ||
                proffixData.FlightBoxData.Immatriculation == "HBZNM")
            {
                // Helicopter over 78dB --> double landing tax price

                if (proffixData.SchHome > 0) proffixData.SchHome = proffixData.SchHome * 2;
                if (proffixData.SchExternal > 0) proffixData.SchExternal = proffixData.SchExternal * 2;
                if (proffixData.LdgTaxHomebased > 0) proffixData.LdgTaxHomebased = proffixData.LdgTaxHomebased * 2;
                if (proffixData.LdgTaxExternal > 0) proffixData.LdgTaxExternal = proffixData.LdgTaxExternal * 2;
            }

            //sum up all landing tax fees
            proffixData.ArticlePrice = proffixData.SchHome + proffixData.SchExternal + proffixData.LdgTaxHomebased +
                                       proffixData.LdgTaxExternal;
        }

        private void WriteBaseFile(List<ProffixData> proffixDataList)
        {
            var folder = ExportFolderName + CreationTimeStamp.ToString("yyyy-MM-dd") + "\\";

            if (Directory.Exists(folder) == false)
            {
                Directory.CreateDirectory(folder);
            }

            var exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_Base.csv");
            var nrOfLinesExported = 0;

            OnLogEventRaised($"Exportiere Basis-Daten in Datei: {exportFilename}");

            using (var w = new StreamWriter(exportFilename))
            {
                var header =
                    "ARP,TYPMO,ACREG,TYPTR,NUMMO,ORIDE,PAX,DATMO,TIMMO,PIMO,TYPPI,DIRDE,CID,CDT,CDM,KEY,Mitgliedernummer,LASTNAME,MTOW,CLUB,HOME_BASE,ORIGINAL_ORIDE,REMARKS";
                w.WriteLine(header);

                foreach (var proffixData in proffixDataList)
                {
                    var sb = new StringBuilder();
                    sb.Append(proffixData.FlightBoxData.Airport);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.MovementType);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Immatriculation);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.TypeOfTraffic);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.NrOfMovements);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Location);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.NrOfPassengers);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.MovementDate);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.MovementTime);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Runway);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.TypePi);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.DirectionOfDeparture);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.CID);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.CreationDate);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.CreationTime);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Key);
                    sb.Append(",");
                    sb.Append(proffixData.MemberNumber); // using the new mapped member number
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Lastname);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.MaxTakeOffWeight);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Club);
                    sb.Append(",");
                    if (proffixData.FlightBoxData.IsHomebased)
                    {
                        sb.Append("1");
                    }
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.OriginalLocation);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Remarks);
                    w.WriteLine(sb.ToString());
                    nrOfLinesExported++;
                }

                w.Flush();
            }

            OnLogEventRaised($"{nrOfLinesExported} Basis-Daten erfolgreich exportiert.");
        }

        private void WriteLogFile()
        {
            try
            {
                var folder = ExportFolderName + CreationTimeStamp.ToString("yyyy-MM-dd") + "\\";

                if (Directory.Exists(folder) == false)
                {
                    Directory.CreateDirectory(folder);
                }

                var exportFilename = Path.Combine(folder, $"{CreationTimeStamp.ToString("yyyy-MM-dd-HHmm")}_Log.log");
                
                using (var w = new StreamWriter(exportFilename))
                {
                    foreach (var logEntry in _logEntries)
                    {
                        w.WriteLine(logEntry);
                    }

                    w.Flush();
                }                                               
            }
            catch (Exception e)
            {
                OnLogEventRaised($"Konnte Logdatei nicht schreiben. Fehlermeldung: {e.Message}");
            }
        }

        private List<FlightBoxData> ReadFile()
        {
            if (File.Exists(ImportFileName) == false)
            {
                throw new FileNotFoundException($"Datei {ImportFileName} existiert nicht!");
            }

            OnLogEventRaised($"Lese Datei: {ImportFileName}...");

            var flightBoxDataList = new List<FlightBoxData>();
            FlightBoxData.ResetCurrentDataRecordId();

            using (var reader = new StreamReader(ImportFileName, Encoding.UTF8))
            {
                var lineNr = 0;
                var headLines = 0;
                var errorLines = 0;
                var additionalLines = 0;
                FlightBoxData lastLineFlightBoxData = null;

                while (reader.EndOfStream == false)
                {
                    var line = reader.ReadLine();
                    lineNr++;

                    if (line == null)
                    {
                        errorLines++;
                        OnLogEventRaised($"Fehlerhafte Zeile gefunden (Zeile = null) bei Zeile {lineNr}");
                        continue;
                    }

                    if (string.IsNullOrEmpty(line))
                    {
                        errorLines++;
                        OnLogEventRaised($"Leere Zeile gefunden bei Zeile {lineNr}");
                        continue;
                    }

                    if (line.StartsWith("ARP"))
                    {
                        headLines++;
                        continue;
                    }

                    var values = line.Split(',');

                    if (values.Length < 23)
                    {
                        if (values.Length == 1 && values.Last().EndsWith("\"") && lastLineFlightBoxData.Remarks.StartsWith("\""))
                        {
                            lastLineFlightBoxData.Remarks += values.Last();
                            additionalLines++;
                            OnLogEventRaised($"Zeile {lineNr} scheint ein zusätzlicher Kommentar zu sein. Zeileninhalt: {line}");
                            continue;
                        }

                        errorLines++;
                        OnLogEventRaised($"Fehlerhafte Zeile {lineNr} kann nicht verarbeitet werden. Zeileninhalt: {line}");
                        continue;
                    }
                    else if (values.Length > 23)
                    {
                        //unescape remarks with commas
                        for (int i = 23; i < values.Length; i++)
                        {
                            values[22] += $",{values[i]}";
                        }

                        values[22] = values[22].Trim('"');
                    }

                    var flightBoxData = new FlightBoxData();

                    try
                    {
                        flightBoxData.LineNumber = lineNr;
                        flightBoxData.Airport = values[0];
                        flightBoxData.MovementType = values[1];
                        flightBoxData.Immatriculation = values[2];
                        flightBoxData.TypeOfTraffic = System.Convert.ToInt32(values[3]);
                        flightBoxData.NrOfMovements = System.Convert.ToInt32(values[4]);
                        flightBoxData.Location = values[5];
                        flightBoxData.NrOfPassengers = System.Convert.ToInt32(values[6]);
                        flightBoxData.MovementDate = values[7];
                        flightBoxData.MovementTime = values[8];

                        DateTime parsedDate;
                        if (DateTime.TryParseExact(flightBoxData.MovementDate, "yyyyMMdd", null, DateTimeStyles.None, out parsedDate))
                        {
                            flightBoxData.MovementDateTime = parsedDate;
                        }
                        else
                        {
                            OnLogEventRaised($"Warnung beim Konvertieren des Datums auf Zeile {lineNr}. Zu konvertierendes Datum: {flightBoxData.MovementDate}");
                        }

                        DateTime parsedTime;
                        if (DateTime.TryParseExact(flightBoxData.MovementTime, "Hmm", null, DateTimeStyles.None, out parsedTime))
                        {
                            flightBoxData.MovementDateTime = flightBoxData.MovementDateTime.AddHours(parsedTime.Hour).AddMinutes(parsedTime.Minute);
                        }
                        else
                        {
                            OnLogEventRaised($"Warnung beim Konvertieren der Zeit auf Zeile {lineNr}. Zu konvertierende Zeit: {flightBoxData.MovementTime}");
                        }

                        flightBoxData.Runway = values[9];
                        flightBoxData.TypePi = values[10];
                        flightBoxData.DirectionOfDeparture = values[11];
                        flightBoxData.CID = values[12];
                        flightBoxData.CreationDate = values[13];
                        flightBoxData.CreationTime = values[14];

                        if (DateTime.TryParseExact(flightBoxData.CreationDate, "yyyyMMdd", null, DateTimeStyles.None, out parsedDate))
                        {
                            flightBoxData.CreationDateTime = parsedDate;
                        }
                        else
                        {
                            OnLogEventRaised($"Warnung beim Konvertieren des Datums auf Zeile {lineNr}. Zu konvertierendes Datum: {flightBoxData.CreationDate}");
                        }

                        if (DateTime.TryParseExact(flightBoxData.CreationTime, "Hmm", null, DateTimeStyles.None, out parsedTime))
                        {
                            flightBoxData.CreationDateTime = flightBoxData.CreationDateTime.AddHours(parsedTime.Hour).AddMinutes(parsedTime.Minute);
                        }
                        else
                        {
                            OnLogEventRaised($"Warnung beim Konvertieren der Zeit auf Zeile {lineNr}. Zu konvertierende Zeit: {flightBoxData.CreationTime}");
                        }

                        flightBoxData.Key = values[15];
                        flightBoxData.MemberNumber = values[16];
                        flightBoxData.Lastname = values[17];

                        if (string.IsNullOrWhiteSpace(flightBoxData.Lastname))
                        {
                            errorLines++;
                            OnLogEventRaised($"Fehlerhafte Zeile {lineNr}. Kein Nachname in Zeile vorhanden. {line}");
                            continue;
                        }

                        flightBoxData.MaxTakeOffWeight = System.Convert.ToInt32(values[18]);
                        flightBoxData.Club = values[19];
                        flightBoxData.IsHomebased = values[20] == "1";
                        flightBoxData.OriginalLocation = values[21];
                        flightBoxData.Remarks = values[22];
                        flightBoxData.SetNextDataRecordId();
                    }
                    catch (Exception e)
                    {
                        errorLines++;
                        OnLogEventRaised($"Fehler beim Konvertieren der Zeile {lineNr}. Message: {e.Message}");
                    }

                    flightBoxDataList.Add(flightBoxData);
                    lastLineFlightBoxData = flightBoxData;
                }

                OnLogEventRaised($"Import durchgeführt. {lineNr} Zeilen eingelesen. Davon {headLines} Kopfzeilen, {additionalLines} Zusatzlinien (Kommentare), {errorLines} fehlerhafte Zeilen ergibt {flightBoxDataList.Count} Datensätze.");
            }

            return flightBoxDataList;
        }
        
        private void OnLogEventRaised(string text)
        {
            _logEntries.Add(text);
            LogEventRaised?.Invoke(this, new LogEventArgs(text));
        }
    }
}
