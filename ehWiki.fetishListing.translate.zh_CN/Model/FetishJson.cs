using Chsword.Excel2Object;

namespace ehWiki.fetishListing.translate.zh_CN.Model
{
    public class FetishJson
    {
        /// <summary>
        /// 父子名称key
        /// </summary>
        [ExcelTitle("父子Key")]
        public string ps_en { get; set; }
        /// <summary>
        /// 父EN
        /// </summary>
        [ExcelTitle("父级EN")]
        public string parent_en { get; set; }
        /// <summary>
        /// 父ZH
        /// </summary>
        [ExcelTitle("父级ZH")]
        public string parent_zh { get; set; }
        /// <summary>
        /// 子EN
        /// </summary>
        [ExcelTitle("子级EN")]
        public string sub_en { get; set; }
        /// <summary>
        /// 子ZH
        /// </summary>
        [ExcelTitle("子级ZH")]
        public string sub_zh { get; set; }
        /// <summary>
        /// 子描述
        /// </summary>
        [ExcelTitle("子级描述")]
        public string sub_desc { get; set; }
    }
}
