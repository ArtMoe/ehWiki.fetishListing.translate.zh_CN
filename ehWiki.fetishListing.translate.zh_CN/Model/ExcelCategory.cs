using Chsword.Excel2Object;

namespace ehWiki.fetishListing.translate.zh_CN.Model
{
    public class ExcelCategory
    {
        [ExcelTitle("父级")]
        public string ParentEn { get; set; }
        [ExcelTitle("父级名称")]
        public string ParentZh { get; set; }
        [ExcelTitle("子级")]
        public string SubEn { get; set; }
        [ExcelTitle("子级名称")]
        public string SubZh { get; set; }
        [ExcelTitle("子级说明")]
        public string SubDesc { get; set; }
    }
}
