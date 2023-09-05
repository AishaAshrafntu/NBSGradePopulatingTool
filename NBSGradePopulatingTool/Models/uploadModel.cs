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
       
        
        [Required(ErrorMessage = "Please select NOW extract file")]
        public required IFormFile NowCsvFile { get; set; }
        [Required(ErrorMessage = "Please select Banner EMS file")]
        public required IFormFile BannerXlsxFile { get; set; }

        
    }

    

    public class ReponseModel
    {
        public string? Message { get; set; }
        public bool IsSuccess { get; set; }
        public bool IsResponse { get; set; }
    }



}
