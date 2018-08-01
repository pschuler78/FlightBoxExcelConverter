﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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


        public DataManager()
        {
            var file = Settings.Default.MemberListFileName;

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

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("Lastname,")) continue;

                        var values = line.Split(',');

                        if (values.Length < 2) continue;

                        _memberList.Add(values[0], values[1]);
                    }
                }
            }

            file = Settings.Default.NoLdgTaxMembersFileName;

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

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("MemberNumber,"))
                            continue;

                        var values = line.Split(',');

                        if (values.Length < 2) continue;

                        _noLdgTaxMembers.Add(values[0]);
                    }
                }
            }
            

            file = Settings.Default.MemberNumberRemappingFileName;

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

                        if (line == null || string.IsNullOrEmpty(line) || line.StartsWith("Immatriculation,"))
                            continue;

                        var values = line.Split(',');

                        if (values.Length < 2) continue;

                        _memberNrRemapping.Add(values[0], values[1]);
                    }
                }
            }
        }

        public bool FindLastnameAndSetMemberNumber(ProffixData proffixData)
        {
            foreach (var lastname in _memberList.Keys)
            {
                if (lastname.ToLower() == proffixData.FlightBoxData.Lastname.ToLower())
                {
                    proffixData.MemberNumber = _memberList[lastname];
                    return true;
                }
            }

            return false;
        }

        public bool IsNoLdgTaxMember(ProffixData proffixData)
        {
            if (_noLdgTaxMembers.Exists(x => x.Contains(proffixData.MemberNumber)))
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