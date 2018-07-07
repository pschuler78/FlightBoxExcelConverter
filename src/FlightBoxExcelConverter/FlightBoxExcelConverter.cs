using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using FlightBoxExcelConverter.Enums;
using FlightBoxExcelConverter.Objects;

namespace FlightBoxExcelConverter
{
    public class FlightBoxExcelConverter
    {
        public event EventHandler<LogEventArgs> LogEventRaised;
        public event EventHandler ExportFinished;

        public string ImportFileName { get; set; }

        public string ExportFileName { get; set; }

        public bool HasExportError { get; set; }

        public string ExportErrorMessage { get; set; }

        public DataCleaner _dataCleaner;

        public DataRemapper _dataRemapper;

        public FlightBoxExcelConverter(string importFileName, string exportFileName)
        {
            ImportFileName = importFileName;
            ExportFileName = exportFileName;
            _dataCleaner = new DataCleaner();
            _dataRemapper = new DataRemapper();
        }

        public void Convert()
        {
            try
            {
                var flightBoxDataList = ReadFile();
                var proffixDataList = new List<ProffixData>();

                foreach (var flightBoxData in flightBoxDataList)
                {
                    CompleteMemberNumbers(flightBoxData);
                }

                foreach (var flightBoxData in flightBoxDataList)
                {
                    CleanupData(flightBoxData);
                }

                foreach (var flightBoxData in flightBoxDataList)
                {
                    var proffixData = ConvertData(flightBoxData);
                    if (proffixData == null)
                    {
                        OnLogEventRaised($"Fehler beim Erstellen der Proffix Daten. Daten in Zeile {flightBoxData.LineNumber} sind fehlerhaft.");
                        continue;
                    }

                    proffixDataList.Add(proffixData);
                }

                OnLogEventRaised($"Exportiere Daten in Datei: {ExportFileName}");
                Thread.Sleep(50);

                WriteFile(proffixDataList);
                OnLogEventRaised($"Daten erfolgreich exportiert.");
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

        private void WriteFile(List<ProffixData> proffixDataList)
        {
            using (var w = new StreamWriter(ExportFileName))
            {
                var header =
                    "ARP,TYPMO,ACREG,TYPTR,NUMMO,ORIDE,PAX,DATMO,TIMMO,PIMO,TYPPI,DIRDE,CID,CDT,CDM,KEY,Mitgliedernummer,LASTNAME,MTOW,CLUB,HOME_BASE,ORIGINAL_ORIDE,,Mitgliedernummer,ArtikelNr,ArtMenge,ArtPreis,VFSArtikelNr,VFSMenge,VFSPreis,SchSpeck,SchFremd,HB,Fremd";
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
                    if (proffixData.FlightBoxData.IsHomebased)
                    {
                        sb.Append("1");
                    }
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.OriginalLocation);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Remarks);
                    sb.Append(",");
                    sb.Append(proffixData.MemberNumber);
                    sb.Append(",");
                    sb.Append(proffixData.ArticleNr);
                    sb.Append(",");
                    sb.Append(proffixData.ArticleQuantity.ToString("0"));
                    sb.Append(",");
                    sb.Append(proffixData.ArticlePrice.ToString("0.00"));
                    sb.Append(",");
                    sb.Append(proffixData.VfsArticleNumber);
                    sb.Append(",");
                    sb.Append(proffixData.VfsQuantity.ToString("0"));
                    sb.Append(",");
                    sb.Append(proffixData.VfsPrice.ToString("0.00"));
                    sb.Append(",");
                    sb.Append(proffixData.SchHome.ToString("0.00"));
                    sb.Append(",");
                    sb.Append(proffixData.SchExternal.ToString("0.00"));
                    sb.Append(",");
                    sb.Append(proffixData.LdgTaxHomebased.ToString("0.00"));
                    sb.Append(",");
                    sb.Append(proffixData.LdgTaxExternal.ToString("0.00"));
                    w.WriteLine(sb.ToString());
                }

                w.Flush();
            }
        }

        private ProffixData ConvertData(FlightBoxData flightBoxData)
        {
            var proffixData = new ProffixData(flightBoxData);

            // set MemberNumber in Proffix data
            if (_dataRemapper.FindImmatruculationAndMapMemberNumber(proffixData))
            {
                OnLogEventRaised($"Setze spezielle Mitgliedernummer für {proffixData.FlightBoxData.Immatriculation} (Zeile: {flightBoxData.LineNumber}): Alte Mitgliedernummer {proffixData.FlightBoxData.MemberNumber}, neue Mitgliedernummer {proffixData.MemberNumber}");
            }
            else
            {
                proffixData.MemberNumber = proffixData.FlightBoxData.MemberNumber;
            }

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
            if (proffixData.FlightBoxData.TypeOfTraffic == (int) TypeOfTraffic.Instruction
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
                && proffixData.FlightBoxData.TypeOfTraffic != (int) TypeOfTraffic.Instruction)
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
                    proffixData.LdgTaxHomebased = 17.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1001)
                {
                    proffixData.LdgTaxHomebased = 20.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1251)
                {
                    proffixData.LdgTaxHomebased = 22.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 1501)
                {
                    proffixData.LdgTaxHomebased = 25.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight < 2001)
                {
                    proffixData.LdgTaxHomebased = 30.0m;
                }
                else if (proffixData.FlightBoxData.MaxTakeOffWeight > 2000)
                {
                    proffixData.LdgTaxHomebased = 40.0m;
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

            //sum up all landing tax fees
            proffixData.ArticlePrice = proffixData.SchHome + proffixData.SchExternal + proffixData.LdgTaxHomebased +
                                       proffixData.LdgTaxExternal;

            return proffixData;
        }

        private void CleanupData(FlightBoxData flightBoxData)
        {
            
        }

        private void CompleteMemberNumbers(FlightBoxData flightBoxData)
        {
            if (string.IsNullOrEmpty(flightBoxData.MemberNumber))
            {
                if (_dataCleaner.FindLastnameAndAddMemberNumber(flightBoxData))
                {
                    OnLogEventRaised($"MemberNumber {flightBoxData.MemberNumber} für {flightBoxData.Lastname} mit {flightBoxData.Immatriculation} gesetzt (Zeile: {flightBoxData.LineNumber}).");
                }
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
                        errorLines++;
                        OnLogEventRaised($"Fehlerhafte Zeile {lineNr} kann nicht verarbeitet werden. Zeileninhalt: {line}");
                        continue;
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
                }

                OnLogEventRaised($"Import durchgeführt. {lineNr} Zeilen eingelesen. Davon {headLines} Kopfzeilen, {errorLines} fehlerhafte Zeilen ergibt {flightBoxDataList.Count} Datensätze.");
            }

            return flightBoxDataList;
        }

        private void OnLogEventRaised(string text)
        {
            LogEventRaised?.Invoke(this, new LogEventArgs(text));
        }
    }
}
