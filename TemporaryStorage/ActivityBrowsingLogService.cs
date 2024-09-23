using Microsoft.EntityFrameworkCore;
using Models.Admin.Ajax;
using Models.Admin.View.ActivityBrowsingLog;
using Models.DBModels;
using Models.Enums;
using Models.Extensions;
using Models.Resources;
using Models.ResponseModels;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Services.GeneralServices;
using Services.Interfaces.AdminServices;
using X.PagedList;

namespace Services.AdminServices
{
    /// <summary>
    /// 活動瀏覽紀錄
    /// </summary>
    public class ActivityBrowsingLogService : BaseService, IActivityBrowsingLogService
    {
        public ActivityBrowsingLogService(
            IConfiguration configuration,
            CmpAppContext db,
            IHttpContextAccessor httpContextAccessor)
        : base(configuration, db, httpContextAccessor)
        {
        }

        /// <summary>
        /// 取得活動瀏覽紀錄列表
        /// </summary>
        /// <param name="query">搜尋參數</param>
        /// <returns></returns>
        public async Task<PagedListResultResponse<ActivityBrowsingLogModel>> GetActivityBrowsingLogs(ActivityBrowsingLogListQueryModel query)
        {
            const string periodFormat = "{0}  ~  {1}";
            const string dateFormat = "yyyy/MM/dd HH:mm";

            if (query.StartAt.HasValue //如果有開始時間
                && query.EndAt.HasValue // 如果有結束時間
                && query.StartAt.Value.IsAfter(query.EndAt.Value)) // 如果開始時間在結束時間之後
            {
                // 回報錯誤
                return new() 
                {
                    Success = false,
                    Code = ResponseCodeMessage.InvalidArgumentsCode.ToString(),
                    Message = ResponseCodeMessage.AdminInvalidDatePeriod
                };
            }

            // activity query // 找活動的部分
            var activitiesQuery = DB.Activities // 找DB的 活動 表
                .Where(a => !a.DeletedAt.HasValue // 沒有被刪除的
                    && (string.IsNullOrEmpty(query.ActivityName) || a.ActivityName.Contains(query.ActivityName)) // 而且 (沒有查詢活動名稱 或是 活動名稱包含查詢的活動名稱)
                    && (!query.StartAt.HasValue || a.StartAt >= query.StartAt.Value) // 而且 (沒有查詢活動開始時間 或是 活動表內的項目的開始時間在查詢的開始時間之後>也就是檢查查詢的活動不是未來的活動)
                    && (!query.EndAt.HasValue || a.EndAt < query.EndAt.Value.AddDays(1)) // 而且 (沒有查詢活動結束時間 或是 活動表內的項目的結束時間在查詢的活動結束時間+1天之前>也就是查有沒有項目的活動結束時間在查詢的活動結束時間後一天之前)
                    && (!query.PlaceId.HasValue || a.PlaceId == query.PlaceId.Value))// 而且 (沒有查詢勤美物業Id 或是 有對應的勤美物業Id)
                .Select(a => new // 顯示出
                {
                    Id = a.Id, // 勤美物業Id
                    ActivityName = a.ActivityName, // 活動名稱
                    ActivityPeriod = string.Format(periodFormat, a.StartAt.ToString(dateFormat), a.EndAt.ToString(dateFormat)), // 活動期間
                    PlaceName = a.Place != null ? a.Place.PlaceName : null, // 如果該項目有 所屬勤美物業 就填入該項名稱，沒有就null
                    ThemeCategory = DB.ThemeCategoryEntities // 主題分類是查詢DB的主題分類實體資料表
                        .Where(tce => tce.EntityType == (int)ThemeCategoryEntityTypeEnum.Activity && tce.EntityId == a.Id) // 他的分類要是 主題分類Enum的 活動 而且他的Id要是查出來活動的項目的Id
                        .Select(tce => new { tce.ThemeCategory.Name, tce.ThemeCategoryId }) // 顯示 分類名稱 和 分類Id
                        .FirstOrDefault(), // 分類找符合條件的第一項
                    CreatedAt = a.CreatedAt  // for sorting // 建立時間
                })
                .Where(a => a.ThemeCategory != null // 找出來的這個Model，還要去找他有 分類 而且(沒有查 分類Id 或是項目的 分類Id 等於查詢的 分類Id)的項目
                    && (!query.ThemeCategoryId.HasValue || a.ThemeCategory.ThemeCategoryId == query.ThemeCategoryId.Value));

            // grouped browsing log query
            var browsingLogsQuery = DB.ActivityBrowsingLogs // 查詢活動瀏覽log
                .GroupBy(log => log.ActivityId) // 以活動Id分組
                .Select(group => new // 每個組顯示以下資訊
                {
                    ActivityId = group.Key, // 活動Id 是那個分組的Key，也就是ActivityId
                    ClickCount = group.Count(), // Count()返回裡面有多少子集合，在這邊就是那個活動有多少瀏覽紀錄
                    MemberCount = group.Select(log => log.UTKMemberId).Distinct().Count() // 遍歷那個活動的瀏覽紀錄抓出那條瀏覽紀錄的 環友會員編號 ，排除重複的，計算它的數量
                });

            PagedListResultResponse<ActivityBrowsingLogModel> result = new() // 返回分頁的活動瀏覽紀錄Model
            {
                Data = new PagedData<ActivityBrowsingLogModel>() // 裡面的Data屬性的類型是 PagedData<T>? 這邊T是活動瀏覽紀錄Model
                {
                    // activity left join browsing log
                    DataList = await activitiesQuery // PagedData<T> 裡面的 DataList屬性是一個 IPagedList<T> 這邊T是ActivityBrowsingLogModel
                        .GroupJoin( // 返回分組結構 主集合元素及底下的子集合
                            browsingLogsQuery, // 內部集合 活動瀏覽紀錄
                            activity => activity.Id, // 外部集合的鍵
                            browsingLog => browsingLog.ActivityId, // 內部集合的鍵
                            (activity, browsingLogs) => new { Activity = activity, BrowsingLogs = browsingLogs.DefaultIfEmpty() }) // 如何顯示: 主集合 活動 子集合 活動瀏覽紀錄，如果為空使用預設值
                        .OrderByDescending(a => a.Activity.CreatedAt) // 以降序排序 套用到時間上 => 從新到舊
                        .SelectMany( // 將查詢出來的內容扁平化 可能是一個陣列
                            a => a.BrowsingLogs,
                            (a, browsingLog) => new ActivityBrowsingLogModel
                            {
                                ActivityId = a.Activity.Id,
                                ActivityName = a.Activity.ActivityName,
                                ActivityPeriod = a.Activity.ActivityPeriod,
                                PlaceName = a.Activity.PlaceName != null ? a.Activity.PlaceName : null,
                                ThemeCategoryName = a.Activity.ThemeCategory!.Name,
                                ClickCount = browsingLog.ClickCount != null ? browsingLog.ClickCount : 0,
                                MemberCount = browsingLog.MemberCount != null ? browsingLog.MemberCount : 0
                            }
                        )
                        .ToPagedListAsync(query.Page, query.PageSize)
                }
            };

            return result;
        }

        /// <summary>
        /// 匯出活動瀏覽紀錄
        /// </summary>
        /// <param name="reqModel">輸入參數</param>
        /// <returns></returns>
        public async Task<ResultResponse<ExportActivityBrowsingLogResponseModel>> ExportActivityBrowsingLog(ExportActivityBrowsingLogRequestModel reqModel)
        {
            ResultResponse<ExportActivityBrowsingLogResponseModel> result = new();
            ExportActivityBrowsingLogResponseModel resModel = new();

            var memberClickCounts = await DB.ActivityBrowsingLogs
                .Where(log => log.ActivityId == reqModel.ActivityId
                    && !string.IsNullOrEmpty(log.UTKMemberId))
                // group by {activityId, UTKMemberId}, and count per member click
                .GroupBy(log => new { log.ActivityId, log.UTKMemberId })
                .Select(log => new MemberClickCountModel
                {
                    UTKMemberId = log.Key.UTKMemberId,
                    MemberClickCount = log.Count()
                })
                .ToListAsync();

            // map to excel model
            var browsingLogExcelModel = new ActivityBrowsingLogExcelModel()
            {
                ActivityName = reqModel.ActivityName,
                PlaceName = reqModel.PlaceName,
                ThemeCategoryName = reqModel.ThemeCategoryName,
                ClickCount = reqModel.ClickCount,
                MemberCount = reqModel.MemberCount,
                memberClickCounts = memberClickCounts
            };

            resModel.FileName = GetExportActivityBrowsingLogFileName(reqModel.ActivityName ?? string.Empty);
            resModel.FileContent = GenerateActivityBrowsingLogExcel(browsingLogExcelModel);

            result.Data = resModel;

            return result;
        }

        /// <summary>
        /// 取得匯出活動瀏覽紀錄檔案名稱
        /// </summary>
        /// <param name="activityName">活動名稱</param>
        /// <returns></returns>
        private string GetExportActivityBrowsingLogFileName(string activityName)
        {
            // format: activityName瀏覽紀錄.xlsx
            var fileNameFormat = "{0}瀏覽紀錄.xlsx";
            return string.Format(fileNameFormat, activityName);
        }

        /// <summary>
        /// 產生活動瀏覽紀錄 Excel
        /// </summary>
        /// <param name="browsingLogExcelModel">活動瀏覽紀錄 excel model</param>
        /// <returns></returns>
        private byte[] GenerateActivityBrowsingLogExcel(ActivityBrowsingLogExcelModel browsingLogExcelModel)
        {
            string[] firstPartHeader = { "活動名稱", "店館類別", "活動主題", "總點擊次數", "瀏覽會員數" };
            string[] secondPartHeader = { "memberID", "瀏覽次數" };

            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Activity Browsing Log");

            // Set column width
            for (int i = 0; i < firstPartHeader.Length; i++)
            {
                worksheet.SetColumnWidth(i, 20 * 256);
            }

            /*  first part */
            // Write header row
            var headerRow = worksheet.CreateRow(0);
            for (int i = 0; i < firstPartHeader.Length; i++)
            {
                headerRow.CreateCell(i).SetCellValue(firstPartHeader[i]);
            }
            // Insert data
            var firstPartRow = worksheet.CreateRow(1);
            firstPartRow.CreateCell(0).SetCellValue(browsingLogExcelModel.ActivityName);
            firstPartRow.CreateCell(1).SetCellValue(browsingLogExcelModel.PlaceName);
            firstPartRow.CreateCell(2).SetCellValue(browsingLogExcelModel.ThemeCategoryName);
            firstPartRow.CreateCell(3).SetCellValue(browsingLogExcelModel.ClickCount ?? 0);
            firstPartRow.CreateCell(4).SetCellValue(browsingLogExcelModel.MemberCount ?? 0);

            /*  second part */
            // Write header row
            var secondPartHeaderRow = worksheet.CreateRow(3);
            for (int i = 0; i < secondPartHeader.Length; i++)
            {
                secondPartHeaderRow.CreateCell(i).SetCellValue(secondPartHeader[i]);
            }
            // Insert data
            int rowIndex = 4;
            foreach (var memberClick in browsingLogExcelModel.memberClickCounts)
            {
                var secondPartRow = worksheet.CreateRow(rowIndex);
                secondPartRow.CreateCell(0).SetCellValue(memberClick.UTKMemberId);
                secondPartRow.CreateCell(1).SetCellValue(memberClick.MemberClickCount ?? 0);
                rowIndex++;
            }

            using var memoryStream = new MemoryStream();
            workbook.Write(memoryStream);
            return memoryStream.ToArray();
        }
    }
}
