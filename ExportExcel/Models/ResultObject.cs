using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExportExcel.Models
{
    public class ResultObject<T>
    {
        public int Code { get; set; }
        public string Msg { get; set; }
        public T Data { get; set; }
        public static ResultObject<T> GetResult(int code ,string msg ,T data = default(T))
        {
            return new ResultObject<T>
            {
                Code = code,
                Msg = msg,
                Data = data
            };
        }
    }
}
