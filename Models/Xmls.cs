using System;

namespace EmailReseiver.Models
{
    public class Xmls
    {
        public int Id { get; set; }
        public string? FileName { get; set; }
        public string? FileContent { get; set; }
        public string? UniqueId { get; set; }
        public DateTime InsertDate { get; set; }
    }
}