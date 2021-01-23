using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using FlightBoxExcelConverter.Enums;
using FlightBoxExcelConverter.Objects;
using FlightBoxExcelConverter.Properties;

namespace FlightBoxExcelConverter
{
    public class DataManager
    {
        private Dictionary<string, string> _memberList = new Dictionary<string, string>();
        private List<string> _noLdgTaxMembers = new List<string>();
        private Dictionary<string, string> _memberNrRemapping = new Dictionary<string, string>();
        private List<string> _proffixAddressNumbers = new List<string>();

        public DataManager()
        {
            var file = Settings.Default.MemberListFileName;
            var lineNr = 0;

            if (Path.IsPathRooted(file) == false)
            {
                string directory =
                    Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

                file = Path.Combine(directory, file);
            }

            if (File.Exists(file))
            {
                using (var reader = new StreamReader(file, Encoding.UTF8))
                {
                    while (reader.EndOfStream == false)
                    {
                        var line = reader.ReadLine();
                        lineNr++;

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("Lastname,")) continue;

                        var values = line.Split(',');

                        if (values.Length < 2)
                        {
                            throw new FormatException($"Fehlerhafte Zeile {lineNr} in Konfigurations-Datei: {file}");
                        };

                        _memberList.Add(values[0], values[1]);
                    }
                }
            }

            file = Settings.Default.NoLdgTaxMembersFileName;
            lineNr = 0;

            if (Path.IsPathRooted(file) == false)
            {
                string directory =
                    Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

                file = Path.Combine(directory, file);
            }

            if (File.Exists(file))
            {
                using (var reader = new StreamReader(file, Encoding.UTF8))
                {
                    while (reader.EndOfStream == false)
                    {
                        var line = reader.ReadLine();
                        lineNr++;

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("MemberNumber,"))
                            continue;

                        var values = line.Split(',');

                        if (values.Length < 2)
                        {
                            throw new FormatException($"Fehlerhafte Zeile {lineNr} in Konfigurations-Datei: {file}");
                        };

                        _noLdgTaxMembers.Add(values[0]);
                    }
                }
            }
            

            file = Settings.Default.MemberNumberRemappingFileName;
            lineNr = 0;

            if (Path.IsPathRooted(file) == false)
            {
                string directory =
                    Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

                file = Path.Combine(directory, file);
            }

            if (File.Exists(file))
            {
                using (var reader = new StreamReader(file, Encoding.UTF8))
                {
                    while (reader.EndOfStream == false)
                    {
                        var line = reader.ReadLine();
                        lineNr++;

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("Immatriculation,"))
                            continue;

                        var values = line.Split(',');

                        if (values.Length < 2)
                        {
                            throw new FormatException($"Fehlerhafte Zeile {lineNr} in Konfigurations-Datei: {file}");
                        };

                        _memberNrRemapping.Add(values[0], values[1]);
                    }
                }
            }
        }

        public int ReadProffixDatabase()
        {
            if (Settings.Default.ReadProffixDbData == false)
                return -1;

            string queryString = "SELECT [AdressNrADR] FROM [ADR_Adressen] where Geloescht = 0";
            string connectionString = Settings.Default.ProffixConnectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        _proffixAddressNumbers.Add(reader[0].ToString());
                    }
                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }
            }

            return _proffixAddressNumbers.Count;
        }

        public bool FindLastnameAndSetMemberNumber(ProffixData proffixData)
        {
            foreach (var lastname in _memberList.Keys)
            {
                if (lastname.Trim().ToLower() == proffixData.FlightBoxData.Lastname.Trim().ToLower())
                {
                    proffixData.MemberNumber = _memberList[lastname];
                    return true;
                }
            }

            return false;
        }

        public bool IsNoLdgTaxMember(ProffixData proffixData)
        {
            if (string.IsNullOrWhiteSpace(proffixData.MemberNumber.Trim()))
                return false;

            if (_noLdgTaxMembers.Exists(x => x.Contains(proffixData.MemberNumber.Trim())))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Searches for a mapping immatriculation to member number set and if found a match, it sets the mapped MemberNumber.
        /// </summary>
        /// <param name="proffixData"></param>
        /// <returns>true, if find a matching immatriculation, otherwise false</returns>
        public bool FindImmatriculationAndMapMemberNumber(ProffixData proffixData)
        {
            foreach (var immatriculation in _memberNrRemapping.Keys)
            {
                if (immatriculation.Trim().ToUpper() == proffixData.FlightBoxData.Immatriculation.Trim().ToUpper())
                {
                    proffixData.MemberNumber = _memberNrRemapping[immatriculation];
                    return true;
                }
            }

            return false;
        }

        public bool FindMemberNumberInProffix(ProffixData proffixData)
        {
            if (Settings.Default.ReadProffixDbData == false)
                return true;

            if (_proffixAddressNumbers.Contains(proffixData.MemberNumber.Trim().ToUpper()))
            {
                return true;
            }

            proffixData.MemberNumberInProffixNotFound = true;
            return false;
        }
    }
}
