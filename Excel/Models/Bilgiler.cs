using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Excel.Models
{
    public class Bilgiler
    {
        [Key]
        public int Id { get; set; } //klondeğeri de bunlar bunu name değerini çekiyor
        public string Ad { get; set; }

        public string Soyad { get; set; }
        public string Sinif { get; set; }
        public int Yas { get; set; } //5tane bunları kolonlarını oluşturuyor önce 

        
        
    }
  
  
}
