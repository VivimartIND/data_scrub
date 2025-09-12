const axios = require("axios");
const XLSX = require("xlsx");

const IMAGEKIT_PRIVATE_KEY = "private_VGcK19i8Dvbqq6V9+gKYAPbr9sM=";
const IMAGEKIT_URL_FILES = "https://api.imagekit.io/v1/files";

async function getFilesInFolder(path) {
  console.log(`Starting to fetch files from folder: ${path}`);
  let files = [];
  let skip = 0;
  const limit = 100;

  while (true) {
    console.log(`Fetching files with skip=${skip}, limit=${limit}`);
    try {
      const res = await axios.get(IMAGEKIT_URL_FILES, {
        auth: { username: IMAGEKIT_PRIVATE_KEY, password: "" },
        params: { path, skip, limit },
      });

      console.log(`Received ${res.data.length} files in this batch`);
      if (res.data.length === 0) {
        console.log("No more files to fetch, exiting loop");
        break;
      }

      files = files.concat(res.data);
      skip += limit;
    } catch (error) {
      console.error(`Error fetching files from ${path} at skip=${skip}:`, error.message);
      break;
    }
  }
  console.log(`Total files fetched from ${path}: ${files.length}`);
  return files;
}

(async () => {
  const folderPath = "/product_images/Products";
  console.log("Starting image fetch process...");
  try {
    const files = await getFilesInFolder(folderPath);

    console.log(`Completed fetching ${files.length} files from ${folderPath}`);

    // Extract image details
    const imageData = files.map(file => ({
      Name: file.name,
      URL: file.url,
      FilePath: file.filePath,
      FileType: file.fileType,
    }));

    console.log("Detailed file information:");
    imageData.forEach((img, index) => {
      console.log(
        `${index + 1}. Name: ${img.Name}, Type: ${img.FileType}, URL: ${img.URL}, Path: ${img.FilePath}`
      );
    });

    // Create Excel workbook and worksheet
    console.log("Creating Excel file...");
    const worksheet = XLSX.utils.json_to_sheet(imageData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Images");

    // Customize column headers (optional)
    worksheet["!cols"] = [
      { wch: 30 }, // Name column width
      { wch: 60 }, // URL column width
      { wch: 40 }, // FilePath column width
      { wch: 15 }, // FileType column width
    ];

    // Write to Excel file
    const excelFileName = "image_data.xlsx";
    XLSX.writeFile(workbook, excelFileName);
    console.log(`Image data successfully saved to ${excelFileName}`);
  } catch (error) {
    console.error("Error in main execution:", error.message);
  }
  console.log("Image fetch and Excel creation process completed.");
})();