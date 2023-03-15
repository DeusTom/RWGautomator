using ExcelFunctionLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using ExcelFunctionLibrary.Models.RWGautomator;
using ExcelFunctionLibrary.Models;

namespace KRGautomator
{
    internal class GenerateRoots
    {
        public static List<RWGWritableData> ExpandSentencesExact(List<RWGWritableData> partialSentences, List<RWGWritableData> uniqueWords)
        {
            List<RWGWritableData> newSentences = new();
            foreach (var sentence in partialSentences)
            {
                foreach (var word in uniqueWords)
                {
                    if (sentence.keyword.Contains(word.keyword))
                    {
                        continue;
                    }

                    // Make a copy of the old sentence
                    string newSentence = sentence.keyword;

                    // Add a new word
                    newSentence += " " + word.keyword;

                    newSentence.Trim();
                    newSentences.Add(new RWGWritableData() { keyword = newSentence, frequency = 0, searchVolume = 0 });
                }
            }

            return newSentences;
        }
        public static List<RWGWritableData> ExpandSentencesBroad(List<RWGWritableData> partialSentences, List<RWGWritableData> uniqueWords)
        {
            List<RWGWritableData> newSentences = new();
            for(int i = 0; i < partialSentences.Count; i++)
            {
                for(int j = i+1; j < uniqueWords.Count; j++)
                {
                    if (partialSentences[i].keyword.Contains(uniqueWords[j].keyword))
                    {
                        continue;
                    }

                    // Make a copy of the old sentence
                    string newSentence = partialSentences[i].keyword;

                    // Add a new word
                    newSentence += " " + uniqueWords[j].keyword;

                    newSentence.Trim();
                    newSentences.Add(new RWGWritableData() { keyword=newSentence, frequency = 0, searchVolume = 0 });
                }
            }

            return newSentences;
        }

        public static List<RWGWritableData> GenerateWordCountExact(List<RWGWritableData> uniqueSentences, List<StackedColumns> primaryKeywords)
        {
            foreach (RWGWritableData root in uniqueSentences)
            {
                root.frequency = 0;
                foreach (StackedColumns keyword in primaryKeywords)
                {

                    if (keyword.stackedCellValues[0].Contains(root.keyword))
                    {
                        root.frequency++;
                        root.searchVolume += int.Parse(keyword.stackedCellValues[1], NumberStyles.AllowThousands);
                    }
                }
            }

            return uniqueSentences;
        }

        public static List<RWGWritableData> GenerateWordCountBroad(List<RWGWritableData> uniqueSentences, List<StackedColumns> primaryKeywords)
        {
            foreach (RWGWritableData root in uniqueSentences)
            {
                root.frequency = 0;
                foreach (StackedColumns keyword in primaryKeywords)
                {
                    int currentWordCount = 0;
                    foreach (string word in root.keyword.Split(" "))
                    {
                        foreach (string kword in keyword.stackedCellValues[0].Split(" "))
                        {
                            if (word == kword)
                            {
                                currentWordCount++;
                                break;
                            }
                        }
                    }
                    if (currentWordCount >= root.keyword.Split(" ").Length)
                    {
                        root.frequency++;
                        root.searchVolume += int.Parse(keyword.stackedCellValues[1], NumberStyles.AllowThousands);
                    }
                }
            }

            return uniqueSentences;
        }
    }
}
