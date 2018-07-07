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
    public class DataCleaner
    {
        private Dictionary<string, string> _memberList = new Dictionary<string, string>();
        private List<string> _noLdgTaxMembers = new List<string>();
        private List<string> _noLdgTaxTypes = new List<string>();

        public DataCleaner()
        {
            if (File.Exists(Settings.Default.MemberListFileName))
            {
                using (var reader = new StreamReader(Settings.Default.MemberListFileName, Encoding.UTF8))
                {
                    while (reader.EndOfStream == false)
                    {
                        var line = reader.ReadLine();

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("Lastname,")) continue;

                        var values = line.Split(',');

                        if (values.Length < 2) continue;

                        _memberList.Add(values[0], values[1]);
                    }
                }
            }

            if (File.Exists(Settings.Default.NoLdgTaxMembersFileName))
            {
                using (var reader = new StreamReader(Settings.Default.NoLdgTaxMembersFileName, Encoding.UTF8))
                {
                    while (reader.EndOfStream == false)
                    {
                        var line = reader.ReadLine();

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("MemberNumber,")) continue;

                        var values = line.Split(',');

                        if (values.Length < 2) continue;

                        _noLdgTaxMembers.Add(values[0]);
                    }
                }
            }

            if (File.Exists(Settings.Default.NoLdgTaxTypeTraffic))
            {
                using (var reader = new StreamReader(Settings.Default.NoLdgTaxTypeTraffic, Encoding.UTF8))
                {
                    while (reader.EndOfStream == false)
                    {
                        var line = reader.ReadLine();

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("TYPTR,")) continue;

                        var values = line.Split(',');

                        if (values.Length < 2) continue;

                        _noLdgTaxTypes.Add(values[0]);
                    }
                }
            }
        }

        public bool FindLastnameAndAddMemberNumber(FlightBoxData flightBoxData)
        {
            foreach (var lastname in _memberList.Keys)
            {
                if (lastname.ToLower() == flightBoxData.Lastname.ToLower())
                {
                    flightBoxData.MemberNumber = _memberList[lastname];
                    return true;
                }
            }

            return false;
        }

        public bool ExistsNoLdgTaxMember(FlightBoxData flightBoxData)
        {
            if (_noLdgTaxMembers.Exists(x => x.Contains(flightBoxData.MemberNumber)))
            {
                return true;
            }

            return false;
        }

        public bool ExistsNoLdgTaxTypeTraffic(FlightBoxData flightBoxData)
        {
            if (_noLdgTaxTypes.Exists(x => x.Contains(flightBoxData.MemberNumber)))
            {
                return true;
            }

            return false;
        }
    }
}
