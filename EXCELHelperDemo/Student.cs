using EXCELHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EXCELHelperDemo
{
    public class Student
    {
        [PropertySeq(1)]
        [PropertyColumnName("學號")]
        [Required]
        public int Id { get; set; }

        [PropertySeq(2)]
        [PropertyColumnName("姓名")]
        [Required]
        public string Name { get; set; }

        [PropertySeq(3)]
        [PropertyColumnName("年齡")]
        public int Age { get; set; }

        [PropertySeq(4)]
        [PropertyColumnName("出生日期")]
        public DateTime BirthDate { get; set; }

        [PropertySeq(5)]
        [PropertyColumnName("成績")]
        public double Score { get; set; }
    }
}
