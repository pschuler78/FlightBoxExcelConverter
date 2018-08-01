using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlightBoxExcelConverter.Objects;

namespace FlightBoxExcelConverter.Exporter
{
    public class ProffixDataCsvExporter
    {
        public int NumberOfLinesExported { get; set; }

        private readonly string _exportFilename;
        private readonly List<ProffixData> _proffixDataList;

        public ProffixDataCsvExporter(string exportFilename, List<ProffixData> proffixDataList)
        {
            _exportFilename = exportFilename;
            _proffixDataList = proffixDataList;
        }

        public void Export()
        {
            NumberOfLinesExported = 0;

            using (var w = new StreamWriter(_exportFilename))
            {
                var header =
                    "ARP,TYPMO,ACREG,TYPTR,NUMMO,ORIDE,PAX,DATMO,TIMMO,PIMO,TYPPI,DIRDE,CID,CDT,CDM,KEY,Mitgliedernummer,LASTNAME,MTOW,CLUB,HOME_BASE,ORIGINAL_ORIDE,Mitgliedernummer,ArtikelNr,ArtMenge,ArtPreis,VFSArtikelNr,VFSMenge,VFSPreis,SchSpeck,SchFremd,HB,Fremd";
                w.WriteLine(header);

                foreach (var proffixData in _proffixDataList)
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
                    NumberOfLinesExported++;
                }

                w.Flush();
            }
        }
    }
}
