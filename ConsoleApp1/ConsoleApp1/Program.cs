using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename = "please_saveme.docx";
            // string fileminon = @"C:\Users\John\OneDrive\Air Force\LEAD reflection MFR.docx";
            string word = @"C:\Users\John\Documents\Word\Test.docx";
         //   FileStream fs = new FileStream(@"C:\Users\John\OneDrive\Air Force\LEAD reflection MFR.docx", FileMode.Open);
            var document = DocX.Create(filename);
            var hole = DocX.Load(word);

            //// hole.Save();
            // var numberedList = document.AddList("First List Item.", 0, ListItemType.Numbered);
            // //Add a numbered list starting at 2
            // document.AddListItem(numberedList, "Second List Item.");
            // document.AddListItem(numberedList, "Third list item.");
            // document.AddListItem(numberedList, "First sub list item", 1);

            // document.AddListItem(numberedList, "Nested item.", 2);
            // document.AddListItem(numberedList, "Fourth nested item.");

            // var bulletedList = document.AddList("First Bulleted Item.", 0, ListItemType.Bulleted);
            // document.AddListItem(bulletedList, "Second bullet item");
            // document.AddListItem(bulletedList, "Sub bullet item", 1);
            // document.AddListItem(bulletedList, "Second sub bullet item", 2);
            // document.AddListItem(bulletedList, "Third bullet item");

            // document.InsertList(numberedList);
            // document.InsertList(bulletedList);


            // //var listee = doc.AddList("Whats good", 0, ListItemType.Numbered);
            // //document.AddListItem(listee,"")
            // //oprah.Add("highhh");

            // var lister = document.Lists;
            // var a = lister[0].Items[0].Text;
            // document.InsertParagraph("Hi bro");


            // document.Save();

            var first_par = hole.Paragraphs[2].Text;

            Paragraph sen = hole.Paragraphs[2];
            Formatting form = new Formatting(); 
            form.Highlight = Highlight.yellow;
            form.Bold = true;
            sen.ReplaceText(first_par, first_par, false,
         System.Text.RegularExpressions.RegexOptions.IgnoreCase,
            form, null, MatchFormattingOptions.ExactMatch);


            Process.Start("WINWORD.EXE", word);

            hole.Save();

            //DocX.Load(fs);

            //Console.WriteLine(holding);
        }

    }
}
