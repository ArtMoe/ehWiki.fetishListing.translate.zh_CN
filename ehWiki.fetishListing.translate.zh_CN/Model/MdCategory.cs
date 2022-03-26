using Chsword.Excel2Object;

namespace ehWiki.fetishListing.translate.zh_CN.Model
{
    public class MdCategory
    {
        [ExcelTitle("父级")]
        public string ParentEn { get; set; }
        [ExcelTitle("父级名称")]
        public string ParentZh { get; set; }
        [ExcelTitle("子级")]
        public string SubEn { get; set; }
        [ExcelTitle("子级名称")]
        public string SubZh { get; set; }
        [ExcelTitle("链接")]
        public string Link { get; set; }
    }

    public class MdCategoryEn
    {
        public string ParentEn { get; set; }
        public string SubEn { get; set; }
        public string Link { get; set; }
    }

    public class MdCategoryZh
    {
        public string ParentZh { get; set; }
        public string SubZh { get; set; }
        public string Link { get; set; }
    }
}
