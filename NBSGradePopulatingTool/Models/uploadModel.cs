using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace UploadFile.Models
{
    public class SingleFileModel : ReponseModel
    {
       
        
        [Required(ErrorMessage = "Please select .csv file")]
        public required IFormFile NowCsvFile { get; set; }
        [Required(ErrorMessage = "Please select Xls or .xlsx file")]
        public required IFormFile BannerXlsxFile { get; set; }

        [Required(ErrorMessage = "Please enter destination path")]
        public required string FilePath { get; set; }



    }

    

    public class ReponseModel
    {
        public string? Message { get; set; }
        public bool IsSuccess { get; set; }
        public bool IsResponse { get; set; }
    }



}
