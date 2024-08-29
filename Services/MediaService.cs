using DocumentFormat.OpenXml.Office2010.PowerPoint;
using Microsoft.AspNetCore.Http;

namespace Services;
public class MediaService
{
    /// <summary>
    /// 檔案上傳
    /// </summary>
    /// <param name="file">檔案</param>
    /// <returns></returns>
    //public async Task<MediaModel> Upload(IFormFile file)
    //{
    //    MediaModel result = new();
    //    string originFileName = Path.GetFileNameWithoutExtension(file.FileName);
    //    string fileExtension = Path.GetExtension(file.FileName);
    //    string key = Guid.NewGuid().ToString();
    //    string fileName = Guid.NewGuid().ToString();
    //    long size = file.Length;
    //    string url = string.Empty;

    //    var fileStream = new MemoryStream();

    //    await file.CopyToAsync(fileStream);

    //    fileStream.Seek(0, SeekOrigin.Begin);

    //    switch (_mediaStorageType)
    //    {
    //        case MediaStorageTypeEnum.S3:
    //            url = await S3Upload(fileStream, key, fileName, fileExtension);
    //            break;
    //        case MediaStorageTypeEnum.Local:
    //            url = await LocalUpload(fileStream, key, fileName, fileExtension);
    //            break;
    //    }

    //    Media media = new()
    //    {
    //        OriginalName = originFileName,
    //        MediaKey = key,
    //        Name = fileName,
    //        Ext = fileExtension,
    //        Size = size,
    //        UpdatedAt = DateTime.Now
    //    };

    //    await _db.Medias.AddAsync(media);
    //    await _db.SaveChangesAsync();

    //    return new()
    //    {
    //        Id = media.Id,
    //        Url = url
    //    };
    //}

    /// <summary>
    /// local 檔案上傳
    /// </summary>
    /// <param name="fileStream">檔案</param>
    /// <param name="key">資料夾名稱</param>
    /// <param name="fileName">檔名</param>
    /// <param name="ext">副檔名</param>
    /// <returns></returns>
    //private async Task<string> LocalUpload(Stream fileStream, string key, string fileName, string ext)
    //{
    //    string saveToDirectory = Path.Combine(_webHostEnvironment.WebRootPath, MediaString.MediaFolder, key);
    //    string saveToPath = Path.Combine(saveToDirectory, $"{fileName}{ext}");

    //    Directory.CreateDirectory(saveToDirectory);

    //    var outputFileStream = new FileStream(saveToPath, FileMode.Create, FileAccess.Write);
    //    fileStream.Position = 0;
    //    await fileStream.CopyToAsync(outputFileStream);

    //    return _projectBaseAddress + Path.Combine(MediaString.MediaFolder, key, $"{fileName}{ext}");
    //}
}
