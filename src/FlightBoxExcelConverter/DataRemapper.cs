using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlightBoxExcelConverter.Objects;
using FlightBoxExcelConverter.Properties;

namespace FlightBoxExcelConverter
{
    public class DataRemapper
    {
        private Dictionary<string, string> _memberNrRemapping = new Dictionary<string, string>();

        public DataRemapper()
        {
            if (File.Exists(Settings.Default.MemberNumberRemappingFileName))
            {
                using (var reader = new StreamReader(Settings.Default.MemberNumberRemappingFileName, Encoding.UTF8))
                {
                    while (reader.EndOfStream == false)
                    {
                        var line = reader.ReadLine();

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("Immatriculation,")) continue;

                        var values = line.Split(',');

                        if (values.Length < 2) continue;

                        _memberNrRemapping.Add(values[0], values[1]);
                    }
                }
            }
        }

        /// <summary>
        /// Searches for a mapping immatriculation to member number set and if found a match, it sets the mapped MemberNumber.
        /// </summary>
        /// <param name="proffixData"></param>
        /// <returns>true, if find a matching immatriculation, otherwise false</returns>
        public bool FindImmatruculationAndMapMemberNumber(ProffixData proffixData)
        {
            foreach (var immatriculation in _memberNrRemapping.Keys)
            {
                if (immatriculation.ToUpper() == proffixData.FlightBoxData.Immatriculation.ToUpper())
                {
                    proffixData.MemberNumber = _memberNrRemapping[immatriculation];
                    return true;
                }
            }

            return false;
        }
    }
}
