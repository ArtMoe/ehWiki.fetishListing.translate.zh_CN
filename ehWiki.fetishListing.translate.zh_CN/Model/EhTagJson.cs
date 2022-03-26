using System.Collections.Generic;

namespace ehWiki.fetishListing.translate.zh_CN.Model
{
    public class EhTagJson
    {
        public string repo { get; set; }
        public object head { get; set; }
        public string version { get; set; }
        public List<EhTagJsonData> data { get; set; }
    }

    public class EhTagJsonData
    {
        public string @namespace { get; set; }
        public EhTagJsonDataFrontMatters frontMatters { get; set; }
        public int count { get; set; }
        public Dictionary<string, EhTagJsonDataItem> data { get; set; }
    }

    // frontMatters
    public class EhTagJsonDataFrontMatters
    {
        public string name { get; set; }
        public string description { get; set; }
        public string key { get; set; }

    }

    // sub items
    public class EhTagJsonDataItem
    {
        public string name { get; set; }
        public string intro { get; set; }
        public string links { get; set; }
    }
}
