using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlightBoxExcelConverter.Objects;

namespace FlightBoxExcelConverter.Exporter
{
    public class ReportExporter
    {
        public int NumberOfLinesExported { get; set; }

        private readonly string _exportFilename;
        private readonly List<ProffixData> _proffixDataList;

        public ReportExporter(string exportFilename, List<ProffixData> proffixDataList)
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
                    "Immatriculation,NrLdg,NrOfPAX,Date,Time,Lastname,ArtMenge,ArtPreis,VFSMenge,VFSPreis";
                w.WriteLine(header);

                foreach (var proffixData in _proffixDataList)
                {
                    var sb = new StringBuilder();
                    sb.Append(proffixData.FlightBoxData.MovementType);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Immatriculation);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.NrOfMovements);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.NrOfPassengers);
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.MovementDateTime.Date.ToString("dd.MM.yyyy"));
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.MovementDateTime.ToString("HH:mm"));
                    sb.Append(",");
                    sb.Append(proffixData.FlightBoxData.Lastname);
                    sb.Append(",");
                    sb.Append(proffixData.ArticleQuantity.ToString("0"));
                    sb.Append(",");
                    sb.Append(proffixData.ArticlePrice.ToString("0.00"));
                    sb.Append(",");
                    sb.Append(proffixData.VfsQuantity.ToString("0"));
                    sb.Append(",");
                    sb.Append(proffixData.VfsPrice.ToString("0.00"));
                    w.WriteLine(sb.ToString());
                    NumberOfLinesExported++;
                }

                w.Flush();
            }
        }
    }
}
