using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = new FileInfo(args[0]);
            // Open and read the XlSX file.
            using (var package = new ExcelPackage(file))
            {
                // Get the work book in the file
                var workBook = package.Workbook;
                if (workBook != null)
                {
                    // Get the first worksheet
                    var sheet = workBook.Worksheets.First();
                    var participants = new List<Participant>();
                    for (int i = 2; i < sheet.Dimension.End.Row + 1; i++)
                    {
                        participants.Add(new Participant()
                            {
                                TimeSubmitted = DateTime.ParseExact(sheet.Cells[i, 1].Text, "M-d-yyyy H:mm:ss", CultureInfo.InvariantCulture),
                                Name = sheet.Cells[i, 2].Text,
                                Skill = sheet.Cells[i, 3].Text,
                                DetailedSkills = sheet.Cells[i, 4].Text.Split(new string[] { ", " }, StringSplitOptions.RemoveEmptyEntries),
                            });
                    }
                    participants = participants.GroupBy(p => p.Name).Select(g => g.OrderByDescending(d => d.TimeSubmitted).First()).ToList(); // Remove duplicate entries
                    var groupCount = (int)Math.Ceiling((double)participants.Count / 5);
                    int curGroup = 0;
                    // Create initial ordering
                    foreach (var participant in participants.OrderBy(p => p.Skill))
                    {
                        participant.groupID = curGroup + 1;
                        curGroup = (curGroup + 1) % groupCount;
                    }

                    var groups = participants.GroupBy(p => p.groupID).OrderBy(g => g.Key).ToDictionary(g => g.Key, g => g.ToList());

                    // Balancing the groups (by approximation, since this is actually NP-hard I think), 
                    // by switching one skilled member from the most skilled group with a not skilled member from the least skilled group
                    for (int i = 0; i < 10; i++)
                    {
                        var groupSecondarySkillCount = groups.ToDictionary(g => g.Key, g => g.Value.SelectMany(p => p.DetailedSkills).Count());
                        var minGroup = groupSecondarySkillCount.Aggregate((l, r) => l.Value < r.Value ? l : r).Key;
                        var maxGroup = groupSecondarySkillCount.Aggregate((l, r) => l.Value > r.Value ? l : r).Key;
                        if (minGroup == maxGroup)
                            break;
                        var minParticipant = groups[minGroup].First(f => f.DetailedSkills.Count() == groups[minGroup].Min(p => p.DetailedSkills.Count()));
                        var maxParticipant = groups[maxGroup].First(f => f.DetailedSkills.Count() == groups[maxGroup].Max(p => p.DetailedSkills.Count()));
                        groups[minGroup].Add(maxParticipant); groups[maxGroup].Remove(maxParticipant);
                        groups[maxGroup].Add(minParticipant); groups[minGroup].Remove(minParticipant);
                    }
                    //var groupSecondarySkillCount2 = groups.ToDictionary(g => g.Key, g => g.Value.SelectMany(p => p.DetailedSkills).Count());
                    //Console.WriteLine(String.Join(", ", groupSecondarySkillCount2.Select(g => String.Format("({0},{1})", g.Key, g.Value))));

                    foreach (var group in participants.GroupBy(p => p.groupID).OrderBy(g => g.Key))
                    {
                        Console.WriteLine(String.Format("Groep {0}: {1}", group.Key, String.Join(", ", group.OrderBy(g => g.Name).Select(p => String.Format("{0}", p.Name, p.Skill)).ToArray())));
                    }
                }
            }
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        private class Participant
        {
            public DateTime TimeSubmitted { get; set; }
            public string Name { get; set; }
            public string Skill { get; set; }
            public string[] DetailedSkills { get; set; }
            public int groupID;
        }

    }
}
