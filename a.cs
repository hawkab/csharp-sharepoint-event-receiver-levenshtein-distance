using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Microsoft.SharePoint;

namespace reestrEventReceiver
{
    public struct LDwithID
    {
        public int Distance;
        public string recId;
    }
    public struct namesForSorting
    {
        public int intID;
        public string sourceText;
    }

    public class info
    {
        static int getLevenshteinDistance(string s, string t)
        {
            if (s.Length == 0 || t.Length == 0)
                return int.MaxValue; 

                int n = s.Length;
                int m = t.Length;
                int[,] d = new int[n + 1, m + 1];

                if (n == 0)
                    return m;
                if (m == 0)
                    return n;
                for (int i = 0; i <= n; d[i, 0] = i++)
                { }
                for (int j = 0; j <= m; d[0, j] = j++)
                { }
                for (int i = 1; i <= n; i++)
                {
                    for (int j = 1; j <= m; j++)
                    {
                        int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                        d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                    }
                }
                return d[n, m];
        }

        static int pretreatment(string what, string t)
        {
            int tmp = getLevenshteinDistance(what, t);
            int currentLevenshteinDistance = ((tmp * 100) / what.Length);
            if (currentLevenshteinDistance > 100)
                currentLevenshteinDistance = ((tmp * 100) / t.Length);
            return (100 - currentLevenshteinDistance);
        }

        static List<LDwithID> getMaximumLevenshteinDistance(List<namesForSorting> ss, string t, int countDisplayedRecords = 3)
        {
            List<LDwithID> LevenshteinDistances = new List<LDwithID>();

            int max = 0, hi = 0;
            foreach (namesForSorting s in ss)
            {
                string id = s.intID.ToString(), what = s.sourceText;

                if (!string.IsNullOrEmpty(what))
                {
                    float len = what.Length / float.Parse(t.Length.ToString());
                    if (len < 2.6 && len > 0.39)
                    {
                        int currentLevenshteinDistance = 0;
                        what = handleStringName(what);
                        t = handleStringName(t);
                        if (what != t)
                            currentLevenshteinDistance = pretreatment(what, t);
                        else
                            currentLevenshteinDistance = 100;

                        LevenshteinDistances.Add(new LDwithID() { Distance = currentLevenshteinDistance, recId = id });

                        if (currentLevenshteinDistance > 60)
                            hi += currentLevenshteinDistance;

                        if (max < currentLevenshteinDistance)
                            max = currentLevenshteinDistance;
                        if (hi > 200 & max == 100) return LevenshteinDistances.OrderByDescending(y => y.Distance).Take(countDisplayedRecords).ToList();
                    }
                }
            }

            var result = LevenshteinDistances.OrderByDescending(y => y.Distance);
            if (max<11)
                countDisplayedRecords = 1;
            else if (result.Count() < countDisplayedRecords)
                countDisplayedRecords = result.Count();

            return result.Take(countDisplayedRecords).ToList(); 
        }
        static List<namesForSorting> getNames(SPList currentList, int except, int len)
        {
            List<namesForSorting> result = new List<namesForSorting>();
            
            SPListItemCollection items = currentList.Items;
            foreach (SPListItem item in items)
            {
                if (item.ID != except)
                {
                    string listName = currentList.RootFolder.Name;
                    switch (listName)
                    {
                        case "pub": result.Add(new namesForSorting() { intID = item.ID, sourceText = ""+item["des"] }); break;
                        case "unm": result.Add(new namesForSorting() { intID = item.ID, sourceText = item["des"] + "" + item["desdoc"] }); break;
                        case "License": result.Add(new namesForSorting() { intID = item.ID, sourceText = "" + item["Name"] }); break;
                        case "Scimen": result.Add(new namesForSorting() { intID = item.ID, sourceText = "" + item["thesis"] }); break;
                        case "Contests": result.Add(new namesForSorting() { intID = item.ID, sourceText = item["nameCourse"] + "" + item["year"] + "" + item["result"] }); break;
                        case "Teachnetwork": result.Add(new namesForSorting() { intID = item.ID, sourceText = item["typeOfWork"] + "" + item["levelOfTraining"] + "" + item["directionOfTraining"] }); break;
                        case "MemberExGroups": result.Add(new namesForSorting() { intID = item.ID, sourceText = "" + item["nameExpertGroup"] }); break;
                        case "WorkNetProject": result.Add(new namesForSorting() { intID = item.ID, sourceText = "" + item["nameProject"] }); break;
                        case "Distancemoduls": result.Add(new namesForSorting() { intID = item.ID, sourceText = "" + item["nameModule"] }); break;
                        case "Memberproject": result.Add(new namesForSorting() { intID = item.ID, sourceText = "" + item["nameProject"] }); break;
                        default: break;
                    }
                }
            }
            result.Reverse();
            return result.OrderByDescending(x => x.sourceText.Length > (len - 10) && x.sourceText.Length < (len + 10)).ToList();
        }
        static string handleStringName(string name)
        {
            char[] charInvalidFileChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (char charInvalid in charInvalidFileChars)
                name = name.Replace(charInvalid, new char());
            name = name.ToLower().Replace(".", "").Replace(",", "").Replace("~", "")
                .Replace("a", "а").Replace("e", "е").Replace("o", "о").Replace("y", "у")
                .Replace("x", "х").Replace("t", "т").Replace("h", "н").Replace("m", "м")
                .Replace("p", "р").Replace("c", "с").Replace("u", "и").Replace("k", "к")
                .Replace("b", "в").Replace("  ", " ").Replace("(", "").Replace(")", "").Trim();
            return name;
        }
        static List<LDwithID> findSimilarFiles(SPItemEventProperties properties)
        {
            SPWeb myWeb = properties.OpenWeb();

            List<long> cFiles = new List<long>();
            List<LDwithID> similarResults = new List<LDwithID>();
            SPList currentList = properties.List;
            SPListItem currentItem = properties.ListItem;
            SPAttachmentCollection curFiles = currentItem.Attachments;
            try
            {
                foreach (String attachmentname in curFiles)
                {
                    String attachmentAbsoluteURL = currentItem.Attachments.UrlPrefix + attachmentname;
                    SPFile attachmentFile = myWeb.GetFile(attachmentAbsoluteURL);
                    cFiles.Add(attachmentFile.Length);
                }
                if (cFiles.Count > 0)
                {
                    SPListItemCollection items = properties.List.Items;
                    foreach (SPListItem item in items)
                    {
                        if (item.ID != currentItem.ID)
                        {
                            SPAttachmentCollection files = item.Attachments;
                            if (files.Count > 0)
                            {
                                foreach (String file in files)
                                {
                                    try
                                    {
                                        String attachmentAbsoluteURL = currentItem.Attachments.UrlPrefix + file;
                                        SPFile attachmentFile = myWeb.GetFile(attachmentAbsoluteURL);

                                        foreach (long cFile in cFiles)
                                        {
                                            if (attachmentFile.Length == cFile)
                                                similarResults.Add(new LDwithID { recId = item.ID.ToString(), Distance = 99 });
                                            if (similarResults.Count > 1) return similarResults;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            return similarResults;
        }
        public static void CalculateLevenshteinDistances(SPItemEventProperties properties)
        {
            SPList currentList = properties.List;
            SPListItem currentItem = properties.ListItem;

            currentItem["LevenshteinDistance"] = "Вычисление..."; try { currentItem.SystemUpdate(); } catch { }
            string listName = currentList.RootFolder.Name, what = "0", resultLD = "", resultID = "";
            object importantValue;
            int except = currentItem.ID, id = 0;
            List<namesForSorting> names = new List<namesForSorting>();

            switch (listName)
            {
                case "pub": importantValue = currentItem["des"]; break;
                case "unm": importantValue = currentItem["des"] + "" + currentItem["desdoc"]; break;
                case "License": importantValue = currentItem["Name"]; break;
                case "Scimen": importantValue = currentItem["thesis"]; break;
                case "Contests": importantValue = currentItem["nameCourse"] + "" + currentItem["year"] + "" + currentItem["result"]; break;
                case "Teachnetwork": importantValue = currentItem["typeOfWork"] + "" + currentItem["levelOfTraining"] + "" + currentItem["directionOfTraining"]; break;
                case "MemberExGroups": importantValue = currentItem["nameExpertGroup"]; break;
                case "WorkNetProject": importantValue = currentItem["nameProject"]; break;
                case "Distancemoduls": importantValue = currentItem["nameModule"]; break;
                case "Memberproject": importantValue = currentItem["nameProject"]; break;
                default: importantValue = ""; break;
            }

            try
            {
                if (importantValue == null)
                    currentItem["LevenshteinDistance"] = "Пустая строка";
                else
                {
                    names = currentList.ItemCount > 1 ? info.getNames(currentList, except, importantValue.ToString().Length) : new List<namesForSorting>();

                    // Получение списка расстояний Левенштейна важного значения текущей записи к важным записям остальных записей реестра.
                    List<LDwithID> Levenshteins = info.getMaximumLevenshteinDistance(names, info.handleStringName(importantValue.ToString()));
                    // Получение списка записей, чьи прикреплённые файлы похожи на файлы из текущей записи.
                    List<LDwithID> SimilarFiles = info.findSimilarFiles(properties);
                    if (Levenshteins.Count > 0 || SimilarFiles.Count > 0)
                    {
                        if (SimilarFiles.Count > 0)
                        {
                            // Поиск и удаление пересекающихся записей.
                            foreach (LDwithID simFile in SimilarFiles)
                            {
                                LDwithID exist = Levenshteins.FirstOrDefault(w => w.recId == simFile.recId);
                                if (exist != null)
                                    //if (exist.Distance > simFile.Distance)
                                    Levenshteins.RemoveAll(r => r.recId == exist.recId);
                            }
                            foreach (LDwithID simFile in SimilarFiles)
                                Levenshteins.Add(simFile);
                        }
                        foreach (LDwithID Levenshtein in Levenshteins)
                        {
                            id = Convert.ToInt32(string.IsNullOrEmpty(Levenshtein.recId) ? "0" : Levenshtein.recId); what = Levenshtein.Distance.ToString();
                            resultLD += what + "%; ";
                            resultID += id + ";#" + currentList.Items.GetItemById(id).Title + ";#";
                        }
                        currentItem["LevenshteinDistance"] = resultLD;
                        currentItem["record"] = resultID;
                    }
                    else
                    {
                        currentItem["LevenshteinDistance"] = "Нет записей для сравнения";
                    }
                }
            }
            catch (Exception ex)
            {
                currentItem["LevenshteinDistance"] = "Ошибка:" + ex.Message;
            }
            try { currentItem.SystemUpdate(); }
            catch { }
        }
    }

    public class pub : SPItemEventReceiver
    {
        public override void ItemUpdating(SPItemEventProperties properties)
        {
            info.CalculateLevenshteinDistances(properties);
        }
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            info.CalculateLevenshteinDistances(properties);
        }

        public override void ItemAdded(SPItemEventProperties properties)
        {
            info.CalculateLevenshteinDistances(properties);
        }
    }
}
