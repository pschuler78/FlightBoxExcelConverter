using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LINQtoCSV;

namespace FlightBoxExcelConverter.Objects
{
    public class ProffixData
    {
        public FlightBoxData FlightBoxData { get; }

        public ProffixData(FlightBoxData flightBoxData)
        {
            FlightBoxData = flightBoxData;
            MemberNumber = flightBoxData.MemberNumber;
        }

        /// <summary>
        /// CSV column name: Mitgliedernummer
        /// </summary>
        public string MemberNumber { get; set; }

        /// <summary>
        /// CSV column name: ArtikelNr
        /// </summary>
        public string ArticleNr { get; set; }

        /// <summary>
        /// CSV column name: ArtMenge
        /// </summary>
        public decimal ArticleQuantity { get; set; }

        /// <summary>
        /// CSV column name: ArtPreis
        /// </summary>
        public decimal ArticlePrice { get; set; }

        /// <summary>
        /// CSV column name: VFSArtikelNr
        /// </summary>
        public string VfsArticleNumber { get; set; }

        /// <summary>
        /// CSV column name: VFSMenge
        /// </summary>
        public decimal VfsQuantity { get; set; }

        /// <summary>
        /// CSV column name: VFSPreis
        /// </summary>
        public decimal VfsPrice { get; set; }

        /// <summary>
        /// CSV column name: SchSpeck
        /// </summary>
        public decimal SchHome { get; set; }

        /// <summary>
        /// CSV column name: SchFremd
        /// </summary>
        public decimal SchExternal { get; set; }

        /// <summary>
        /// CSV column name: HB
        /// </summary>
        public decimal LdgTaxHomebased { get; set; }

        /// <summary>
        /// CSV column name: Fremd
        /// </summary>
        public decimal LdgTaxExternal { get; set; }
    }
}
