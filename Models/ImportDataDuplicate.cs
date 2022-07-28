﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace EmailReseiver.Models
{
    public class ImportDataDuplicate
    {
        public ImportDataDuplicate()
        { }

        public ImportDataDuplicate(ImportData baseObj)
        {
            OrgName = baseObj.OrgName;
            MOD = baseObj.MOD;
            INN = baseObj.INN;
            OKPO = baseObj.OKPO;
            ProductName = baseObj.ProductName;
            MedForm = baseObj.MedForm;
            SeriaNum = baseObj.SeriaNum;
            MNN = baseObj.MNN;
            MKB = baseObj.MKB;
            RecSeria = baseObj.RecSeria;
            RecNum = baseObj.RecNum;
            RecDate = baseObj.RecDate;
            OtpuskDate = baseObj.OtpuskDate;
            Quant = baseObj.Quant;
            OkeiName = baseObj.OkeiName;
            Price = baseObj.Price;
            PSum = baseObj.PSum;
            LastName = baseObj.LastName;
            Name = baseObj.Name;
            MidName = baseObj.MidName;
            DateOB = baseObj.DateOB;
            SNILS = baseObj.SNILS;
            InsertDate = baseObj.InsertDate;
            FinancingItem = baseObj.FinancingItem;
            WorkSupplierDogovorId = baseObj.WorkSupplierDogovorId;
        }
        public Int64 Id { get; set; }
        [Required(ErrorMessage = "Не указано имя организации")]
        public string? OrgName { get; set; }
        [Required(ErrorMessage = "Не указан МОД")]
        [StringLength(14, ErrorMessage = "Недоступимая длина МОД")]
        public string? MOD { get; set; }
        [Required(ErrorMessage = "Не указан ИНН")]
        public string INN { get; set; }
        [Required(ErrorMessage = "Не указано ОКПО")]
        public string OKPO { get; set; }
        [Required(ErrorMessage = "Не указано наименование лек-го перепарата")]
        public string? ProductName { get; set; }
        [Required(ErrorMessage = "Не указана лекарств-я форма")]
        public string? MedForm { get; set; }
        [Required]
        public string? SeriaNum { get; set; }
        [Required]
        public string? MNN { get; set; }
        [Required]
        public string MKB { get; set; }
        [Required]
        public string RecSeria { get; set; }
        [Required]
        public string? RecNum { get; set; }
        [Required]
        public DateTime RecDate { get; set; }
        [Required]
        public DateTime OtpuskDate { get; set; }
        [Required]
        [RegularExpression(@"^\d+(\.\d{1,2})?$")]
        [Range(0, 9999999999999999.99)]
        public decimal Quant { get; set; }
        [Required]
        public string? OkeiName { get; set; }
        [Required]
        [RegularExpression(@"^\d+(\.\d{1,2})?$")]
        [Range(0, 9999999999999999.99)]
        public decimal Price { get; set; }
        [Required]
        [RegularExpression(@"^\d+(\.\d{1,2})?$")]
        [Range(0, 9999999999999999.99)]
        public decimal PSum { get; set; }
        [Required(ErrorMessage = "Не указана фамилия пациента")]
        public string? LastName { get; set; }
        [Required(ErrorMessage = "Не указано имя пациента")]
        public string? Name { get; set; }
        [Required(ErrorMessage = "Не указано отчество пациента")]
        public string? MidName { get; set; }
        [Required]
        public DateTime DateOB { get; set; }
        [Required]
        public string? SNILS { get; set; }
        [Required]
        public DateTime InsertDate { get; set; }
        [Required]
        public string FinancingItem { get; set; }
        public string WorkSupplierDogovorId { get; set; }
    }
}
