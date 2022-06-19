using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KRGautomator
{
    internal class GenerateRoots
    {
        public static List<ExcelFunctionLibrary.StringAndInt> ExpandSentencesExact(List<ExcelFunctionLibrary.StringAndInt> partialSentences, List<ExcelFunctionLibrary.StringAndInt> uniqueWords)
        {
            List<ExcelFunctionLibrary.StringAndInt> newSentences = new();
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
                    newSentences.Add(new ExcelFunctionLibrary.StringAndInt() { keyword=newSentence, number=0 });
                }
            }

            return newSentences;
        }
        public static List<ExcelFunctionLibrary.StringAndInt> ExpandSentencesBroad(List<ExcelFunctionLibrary.StringAndInt> partialSentences, List<ExcelFunctionLibrary.StringAndInt> uniqueWords)
        {
            List<ExcelFunctionLibrary.StringAndInt> newSentences = new();
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
                    newSentences.Add(new ExcelFunctionLibrary.StringAndInt() { keyword=newSentence, number = 0});
                }
            }

            return newSentences;
        }

        public static List<ExcelFunctionLibrary.StringAndInt> GenerateWordCountExact(List<ExcelFunctionLibrary.StringAndInt> uniqueSentences, List<string> primaryKeywords)
        {
            foreach (ExcelFunctionLibrary.StringAndInt root in uniqueSentences)
            {
                root.number = 0;
                foreach (string keyword in primaryKeywords)
                {

                    if (keyword.Contains(root.keyword))
                    {
                        root.number++;
                    }
                }
            }

            return uniqueSentences;
        }

        public static List<ExcelFunctionLibrary.StringAndInt> GenerateWordCountBroad(List<ExcelFunctionLibrary.StringAndInt> uniqueSentences, List<string> primaryKeywords)
        {
            foreach (ExcelFunctionLibrary.StringAndInt root in uniqueSentences)
            {
                root.number = 0;
                foreach (string keyword in primaryKeywords)
                {
                    int currentWordCount = 0;
                    foreach (string word in root.keyword.Split(" "))
                    {
                        foreach (string kword in keyword.Split(" "))
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
                        root.number++;
                    }
                }
            }

            return uniqueSentences;
        }
    }
}
