using Chsword.Excel2Object;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ehWiki.fetishListing.translate.zh_CN.Model
{
    /// <summary>
    /// 恋物模块功能
    /// </summary>
    public class FetishListing
    {
        // 根据恋物网页数据，以及EhTag翻译，整合成json数据
        public static void GetFetishListJson()
        {
            // 获取Excel获取原始数据，只需要 sub_en，sub_zh，sub_desc
            var originTranslate = GetOriginExcelData();
            var originParentDict = originTranslate.Item1;
            var originSubDict = originTranslate.Item2;

            // 从 EhTag 中获取 子项数据 sub_en, sub_zh, sub_desc
            Dictionary<string, string[]> ehTagDict = GetEhTagData();

            // 从 markdown 文档中读取匹配中英文，生成初步恋物翻译
            List<MdCategory> markdownTranslate = GetMarkdownData();

            // 匹配 Excel 原始数据，生成恋物数据
            List<FetishJson> fetishList = new List<FetishJson>();
            foreach (var item in markdownTranslate)
            {
                var fetish = new FetishJson();
                if (originParentDict.ContainsKey(item.ParentEn))
                {
                    fetish.parent_zh = originParentDict[item.ParentEn];
                }
                else
                {
                    fetish.parent_zh = item.ParentEn;
                }

                if (originSubDict.ContainsKey(item.SubEn))
                {
                    var originSub = originSubDict[item.SubEn];
                    fetish.sub_zh = originSub[0];
                    fetish.sub_desc = originSub[1];
                }
                else
                {
                    fetish.sub_zh = item.SubZh;
                }
                fetish.parent_en = item.ParentEn;
                fetish.sub_en = item.SubEn;
                fetish.ps_en = $"{fetish.parent_en}:{fetish.sub_en}";
                fetishList.Add(fetish);
            }

            // 导出查看效果
            var fetishListJson = JsonConvert.SerializeObject(fetishList);
            Output("恋物过滤1数据列表", "txt", fetishListJson);

            // 匹配 EhTag 和 恋物，更新符合条件的中文翻译
            foreach (var item in fetishList)
            {
                if (ehTagDict.ContainsKey(item.sub_en) && !string.IsNullOrEmpty(ehTagDict[item.sub_en][0]))
                {
                    var ehTag = ehTagDict[item.sub_en];
                    item.sub_zh = ehTag[0];
                    item.sub_desc = ehTag[1];
                }
            }

            // 导出查看效果
            var fetishListJson2 = JsonConvert.SerializeObject(fetishList);
            Output("恋物过滤2数据列表", "txt", fetishListJson2);

            const string VERSION = "20220423.1309"; // 恋物网页更新日期

            // 生成符合的新版 父子标签json
            var releaseModel = new
            {
                version = VERSION,
                parent_en_array = originParentDict.Select(c => c.Key).ToArray(),
                count = fetishList.Count,
                data = fetishList.ToDictionary(
                    k => k.ps_en,
                    v => new
                    {
                        v.parent_en,
                        v.parent_zh,
                        v.sub_en,
                        v.sub_zh,
                        v.sub_desc
                    })
            };

            // 导出最终发布的数据json [父子同级版本]
            var releaseJson = JsonConvert.SerializeObject(releaseModel);
            ReleaseJson("fetish.oneLevel.json", releaseJson);

            // 导出 Excel 方便检查 [父子同级版本]
            ReleaseExcel("fetish.oneLevel.xlsx", fetishList);

            const string LANG = "Language";
            var fetishWithoutLangList = fetishList.Where(c => c.parent_en != LANG).ToList();
            var releaseWithoutLanguageModel = new
            {
                version = VERSION,
                parent_en_array = originParentDict.Where(c => c.Key != LANG).Select(c => c.Key).ToArray(),
                count = fetishWithoutLangList.Count,
                data = fetishWithoutLangList.ToDictionary(
                    k => k.ps_en,
                    v => new
                    {
                        v.parent_en,
                        v.parent_zh,
                        v.sub_en,
                        v.sub_zh,
                        v.sub_desc
                    })
            };

            // 导出最终发布的数据json [父子同级版本 - 没有 Language 模块]
            var releaseWithOutLangJson = JsonConvert.SerializeObject(releaseWithoutLanguageModel);
            ReleaseJson("fetish.oneLevel.withoutLang.json", releaseWithOutLangJson);

            // 导出 Excel 方便检查 [父子同级版本 - 没有 Language 模块]
            ReleaseExcel("fetish.oneLevel.withoutLang.xlsx", fetishWithoutLangList);

            var fetishWithoutLangSearchKeyList = fetishList.Where(c => c.parent_en != LANG)
                .Select(c => new FetishSearchKeyJson
                {
                    search_key = $"{c.parent_en},{c.parent_zh},{c.sub_en},{c.sub_zh}",
                    parent_en = c.parent_en,
                    parent_zh = c.parent_zh,
                    ps_en = c.ps_en,
                    sub_en = c.sub_en,
                    sub_zh = c.sub_zh,
                    sub_desc = c.sub_desc
                }).ToList();
            var releaseWithoutLanguageSearchKeyModel = new
            {
                version = VERSION,
                parent_en_array = originParentDict.Where(c => c.Key != LANG).Select(c => c.Key).ToArray(),
                count = fetishWithoutLangSearchKeyList.Count,
                data = fetishWithoutLangSearchKeyList.ToDictionary(
                    k => k.ps_en,
                    v => new
                    {
                        v.search_key,
                        v.parent_en,
                        v.parent_zh,
                        v.sub_en,
                        v.sub_zh,
                        v.sub_desc
                    })
            };
            // 导出最终发布的数据json [父子同级版本 - 没有 Language 模块 - 含有搜索关键字]
            var releaseWithOutLangSearchKeyJson = JsonConvert.SerializeObject(releaseWithoutLanguageSearchKeyModel);
            ReleaseJson("fetish.oneLevel.withoutLang.searchKey.json", releaseWithOutLangSearchKeyJson);

            // 导出 Excel 方便检查 [父子同级版本 - 没有 Language 模块]
            ReleaseExcel("fetish.oneLevel.withoutLang.searchKey.xlsx", fetishWithoutLangSearchKeyList);

            var releaseNormalModel = new
            {
                version = VERSION,
                count = fetishList.Count,
                data = originParentDict
                .ToDictionary(k => k.Key, v =>
                {
                    var subItems = fetishList.Where(c => c.parent_en == v.Key).ToList();
                    var parent = new
                    {
                        parent_zh = v.Value,
                        count = subItems.Count,
                        data = subItems
                        .ToDictionary(ks => ks.sub_en, vs =>
                         new
                         {
                             vs.sub_zh,
                             vs.sub_desc
                         })
                    };
                    return parent;
                })
            };

            // 导出最终发布的数据json [父子分级版本]
            var releaseNormalJson = JsonConvert.SerializeObject(releaseNormalModel);
            ReleaseJson("fetish.json", releaseNormalJson);
        }

        // 获取 EhTag 翻译数据
        private static Dictionary<string, string[]> GetEhTagData()
        {
            var tagPath = Path.Combine(Environment.CurrentDirectory, "Json", "db.text.json");
            var tagText = FileTools.Read(tagPath);
            var ehTags = JsonConvert.DeserializeObject<EhTagJson>(tagText);
            Dictionary<string, string[]> ehTagDict = new Dictionary<string, string[]>();
            foreach (var item in ehTags.data)
            {
                var subItemDict = item.data;
                foreach (var subItem in subItemDict)
                {
                    if (!ehTagDict.ContainsKey(subItem.Key) && !string.IsNullOrEmpty(subItem.Value.intro))
                    {
                        ehTagDict[subItem.Key] = new string[] { subItem.Value.name, subItem.Value.intro };
                    }
                }
            }

            // 导出查看效果
            var ehTagDictJson = JsonConvert.SerializeObject(ehTagDict);
            Output("EhTag数据列表", "txt", ehTagDictJson);

            return ehTagDict;
        }

        // 获取原始 Excel 翻译数据
        private static Tuple<Dictionary<string, string>, Dictionary<string, string[]>> GetOriginExcelData()
        {
            var excelPath = Path.Combine(Environment.CurrentDirectory, "Excel", "标签列表.xlsx");
            var importer = new ExcelImporter();
            var excelResult = importer.ExcelToObject<ExcelCategory>(excelPath);
            Dictionary<string, string> originParentDict = new Dictionary<string, string>();
            Dictionary<string, string[]> originSubDict = new Dictionary<string, string[]>();
            foreach (var item in excelResult)
            {
                if (!originParentDict.ContainsKey(item.ParentEn) && !string.IsNullOrEmpty(item.ParentEn))
                {
                    originParentDict[item.ParentEn] = item.ParentZh;
                }

                if (!originSubDict.ContainsKey(item.SubEn))
                {
                    originSubDict[item.SubEn] = new string[] { item.SubZh, item.SubDesc };
                }
            }

            // 导出查看效果
            var originParentJson = JsonConvert.SerializeObject(originParentDict);
            Output("原始数据_父级", "txt", originParentJson);
            var originSubJson = JsonConvert.SerializeObject(originSubDict);
            Output("原始数据_子级", "txt", originSubJson);

            return new Tuple<Dictionary<string, string>, Dictionary<string, string[]>>(originParentDict, originSubDict);
        }

        // 获取恋物网页数据
        private static List<MdCategory> GetMarkdownData()
        {
            var enText = Input("Markdown", "category_en.md");
            var zhText = Input("Markdown", "category_zh.md");
            var splitEmpty = new string[] { "\r\n" };
            string[] enTextArray = enText.Split(splitEmpty, StringSplitOptions.RemoveEmptyEntries);
            var enList = new List<MdCategoryEn>();
            string enTitle = "";
            foreach (var item in enTextArray)
            {
                // 跳过一级标题
                Regex regTitle1 = new Regex("^# (.+)");
                if (regTitle1.IsMatch(item)) continue;

                // 检查是否是二级标题、三级标题、四级标题、五级标题、六级标题
                Regex regTitle2 = new Regex("^[#]{2,} (.+)");
                if (regTitle2.IsMatch(item))
                {
                    var title2 = item.TrimStart('#').Trim();
                    Regex regLink = new Regex("\\((.+)\\)");
                    if (regLink.IsMatch(title2))
                    {
                        title2 = regLink.Replace(title2, "").Replace("[", "").Replace("]", "").Trim(' ', '[', ']', '♂', '♀');
                    }
                    enTitle = title2;
                }
                else
                {
                    // 子项
                    var itemArray = item.Split(',');
                    foreach (var subItem in itemArray)
                    {
                        Regex regLink = new Regex("\\((.+)\\)");
                        var enlink = regLink.Matches(subItem)[0].Value.Trim('(', ')');
                        var enName = regLink.Replace(subItem, "").Replace("[", "").Replace("]", "").Replace("♂", "").Replace("♀", "").Replace("‎", "").Trim();

                        var enItem = new MdCategoryEn
                        {
                            ParentEn = enTitle,
                            SubEn = enName,
                            Link = enlink
                        };
                        enList.Add(enItem);
                    }
                }
            }

            string[] zhTextArray = zhText.Split(splitEmpty, StringSplitOptions.RemoveEmptyEntries);
            var zhList = new List<MdCategoryZh>();
            string zhTitle = "";
            foreach (var item in zhTextArray)
            {
                // 跳过一级标题
                Regex regTitle1 = new Regex("^# (.+)");
                if (regTitle1.IsMatch(item)) continue;

                // 检查是否是二级标题、三级标题、四级标题、五级标题、六级标题
                Regex regTitle2 = new Regex("^[#]{2,} (.+)");
                if (regTitle2.IsMatch(item))
                {
                    var title2 = item.TrimStart('#').Trim();
                    Regex regLink = new Regex("\\((.+)\\)");
                    if (regLink.IsMatch(title2))
                    {
                        title2 = regLink.Replace(title2, "").Replace("[", "").Replace("]", "").Trim(' ', '[', ']', '♂', '♀');
                    }
                    zhTitle = title2;
                }
                else
                {
                    // 子项
                    var itemArray = item.Replace('、', ',').Replace('，', ',').Split(',');
                    foreach (var subItem in itemArray)
                    {
                        Regex regLink = new Regex("\\((.+)\\)");
                        if (regLink.IsMatch(subItem))
                        {
                            var zhlink = regLink.Matches(subItem)[0].Value.Trim('(', ')');
                            var zhName = regLink.Replace(subItem, "").Replace("[", "").Replace("]", "").Replace("♂", "").Replace("♀", "").Trim();

                            var zhItem = new MdCategoryZh
                            {
                                ParentZh = zhTitle,
                                SubZh = zhName,
                                Link = zhlink
                            };
                            zhList.Add(zhItem);
                        }
                        else
                        {
                            var zhItem = new MdCategoryZh
                            {
                                ParentZh = zhTitle,
                                SubZh = subItem
                            };
                            zhList.Add(zhItem);
                        }
                    }
                }
            }

            // 中英文组合
            var mdCategoryList = new List<MdCategory>();
            int index = 0;
            foreach (var enItem in enList)
            {
                var mdcategory = new MdCategory
                {
                    ParentEn = enItem.ParentEn,
                    SubEn = enItem.SubEn,
                    Link = enItem.Link
                };

                var zhItem = zhList.FirstOrDefault(c => c.Link == enItem.Link);
                if (zhItem != null)
                {
                    mdcategory.ParentZh = zhItem.ParentZh;
                    mdcategory.SubZh = zhItem.SubZh;
                }

                mdCategoryList.Add(mdcategory);
                //Console.WriteLine($"{mdcategory.ParentEn}, {mdcategory.ParentZh}, {mdcategory.SubEn}, {mdcategory.SubZh}, {mdcategory.Link}");
                index++;
            }

            // 导出查看效果
            var mdCategoryJson = JsonConvert.SerializeObject(mdCategoryList);
            Output("恋物网页数据列表", "txt", mdCategoryJson);

            return mdCategoryList;
        }


        /// <summary>
        /// 输出文本
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="suffix">后缀名</param>
        /// <param name="data">文本信息</param>
        private static void Output(string fileName, string suffix, string data)
        {
            var baseDict = Path.Combine(Environment.CurrentDirectory, "Output");
            if (!Directory.Exists(baseDict))
            {
                Directory.CreateDirectory(baseDict);
            }

            fileName = $"{fileName}_{DateTime.Now:yyyyMMddHHmmss}.{suffix}";
            var filePath = Path.Combine(baseDict, fileName);
            FileTools.Write(filePath, data);
            Console.WriteLine($"{fileName} 生成完毕!\n");
        }

        /// <summary>
        /// 读取文本
        /// </summary>
        /// <param name="folderName">文件夹名称</param>
        /// <param name="fileNameAndSuffix">文件名称带后缀</param>
        private static string Input(string folderName, string fileNameAndSuffix)
        {
            var dict = Path.Combine(Environment.CurrentDirectory, folderName);
            var path = Path.Combine(dict, fileNameAndSuffix);
            return FileTools.Read(path);
        }

        /// <summary>
        /// 编译新版本json
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="data">数据</param>
        private static void ReleaseJson(string fileName, string data)
        {
            var baseDict = Path.Combine(Environment.CurrentDirectory, "Release_Fetish");
            if (!Directory.Exists(baseDict))
            {
                Directory.CreateDirectory(baseDict);
            }
            var filePath = Path.Combine(baseDict, fileName);
            FileTools.Write(filePath, data);
            Console.WriteLine($"{fileName} 生成完毕!\n");
        }

        /// <summary>
        /// 编译新版本excel
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="data">数据</param>
        private static void ReleaseExcel(string fileName, List<FetishJson> data)
        {
            var baseDict = Path.Combine(Environment.CurrentDirectory, "Release_Fetish");
            if (!Directory.Exists(baseDict))
            {
                Directory.CreateDirectory(baseDict);
            }

            var outputExcelPath = Path.Combine(baseDict, fileName);

            var exporter = new ExcelExporter();
            var bytes = exporter.ObjectToExcelBytes(data, ExcelType.Xlsx);
            File.WriteAllBytes(outputExcelPath, bytes);
            Console.WriteLine($"{fileName} 生成完毕!\n");
        }

        private static void ReleaseExcel(string fileName, List<FetishSearchKeyJson> data)
        {
            var baseDict = Path.Combine(Environment.CurrentDirectory, "Release_Fetish");
            if (!Directory.Exists(baseDict))
            {
                Directory.CreateDirectory(baseDict);
            }

            var outputExcelPath = Path.Combine(baseDict, fileName);

            var exporter = new ExcelExporter();
            var bytes = exporter.ObjectToExcelBytes(data, ExcelType.Xlsx);
            File.WriteAllBytes(outputExcelPath, bytes);
            Console.WriteLine($"{fileName} 生成完毕!\n");
        }
    }
}
