using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Models
{
    /// <summary>
    /// 订金单据model
    /// </summary>
    public class Deposit
    {
        /// <summary>
        /// 药品编号
        /// </summary>
        public string Number { get; set; }
        /// <summary>
        /// 药品名
        /// </summary>
        public string Drugname { get; set; }
        /// <summary>
        /// 数量
        /// </summary>
        public int Quantity { get; set; }
        /// <summary>
        /// 数量单位
        /// </summary>
        public string UnitName { get; set; }
        /// <summary>
        /// 价格
        /// </summary>
        public decimal? Price { get; set; }
        /// <summary>
        /// 订单日期
        /// </summary>
        public DateTime? OrderTime { get; set; }
    }
}
