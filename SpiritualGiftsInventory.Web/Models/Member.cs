using System;

namespace SpiritualGiftsInventory.Web.Models 
{
    public class Member
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Church { get; set; }
        public string SendDate { get; set; }
        public string[] Answers { get; set; }
        public string Punctuation { get; set;}
    }
}