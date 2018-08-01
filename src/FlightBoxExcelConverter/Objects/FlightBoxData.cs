using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlightBoxExcelConverter.Enums;
using LINQtoCSV;

namespace FlightBoxExcelConverter.Objects
{
    public class FlightBoxData
    {
        private static int _currentDataRecordId;

        public void SetNextDataRecordId()
        {
            _currentDataRecordId++;
            DataRecordId = _currentDataRecordId;
        }

        public static void ResetCurrentDataRecordId()
        {
            _currentDataRecordId = 0;
        }

        /// <summary>
        /// Internal line number of data record
        /// </summary>
        public int DataRecordId { get; set; }

        /// <summary>
        /// Original line number of imported record
        /// </summary>
        public int LineNumber { get; set; }

        /// <summary>
        /// ARP
        /// </summary>
        public string Airport { get; set; }

        /// <summary>
        /// TYPMO (Movement types):
        /// V = circuits
        /// A = arrival
        /// D = departure
        /// </summary>
        public string MovementType { get; set; } 

        /// <summary>
        /// ACREG
        /// </summary>
        public string Immatriculation { get; set; }

        /// <summary>
        /// TYPTR
        /// </summary>
        public int TypeOfTraffic { get; set; }

        /// <summary>
        /// NUMMO
        /// </summary>
        public int NrOfMovements { get; set; }

        /// <summary>
        /// ORIDE
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// PAX
        /// </summary>
        public int NrOfPassengers { get; set; }

        /// <summary>
        /// DATMO
        /// </summary>
        public string MovementDate { get; set; }

        /// <summary>
        /// TIMMO
        /// </summary>
        public string MovementTime { get; set; }

        /// <summary>
        /// Combined and parsed movement date and time
        /// </summary>
        public DateTime MovementDateTime { get; set; }

        /// <summary>
        /// PIMO
        /// </summary>
        public string Runway { get; set; }

        /// <summary>
        /// TYPPI: Seams to be always the value 'G'
        /// </summary>
        public string TypePi { get; set; } = "G";

        /// <summary>
        /// DIRDE
        /// </summary>
        public string DirectionOfDeparture { get; set; }

        /// <summary>
        /// CID is same as ARP
        /// </summary>
        public string CID { get; set; }

        /// <summary>
        /// CDT
        /// </summary>
        public string CreationDate { get; set; }

        /// <summary>
        /// CDM
        /// </summary>
        public string CreationTime { get; set; }

        /// <summary>
        /// Combined and parsed creation date and time
        /// </summary>
        public DateTime CreationDateTime { get; set; }

        /// <summary>
        /// KEY
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// MEMBERNR
        /// </summary>
        public string MemberNumber { get; set; }

        /// <summary>
        /// LASTNAME
        /// </summary>
        public string Lastname { get; set; }

        /// <summary>
        /// MTOW
        /// </summary>
        public int MaxTakeOffWeight { get; set; }

        /// <summary>
        /// CLUB
        /// </summary>
        public string Club { get; set; }

        /// <summary>
        /// HOME_BASE
        /// </summary>
        public bool IsHomebased { get; set; }

        /// <summary>
        /// ORIGINAL_ORIDE
        /// </summary>
        public string OriginalLocation { get; set; }

        /// <summary>
        /// REMARKS
        /// </summary>
        public string Remarks { get; set; }

        #region Additional logic properties

        public bool IsTowFlight
        {
            get { return TypeOfTraffic == (int) Enums.TypeOfTraffic.Aerotow; }
        }

        public bool IsMaintenanceFlight { get; set; }

        public bool IsDepartureMovement
        {
            get { return MovementType == "D"; }
        }

        public bool IgnoreLandingTax { get; set; }

        #endregion Additional logic properties
    }
}
