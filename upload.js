const express = require('express');
const multer = require('multer');
const AWS = require('aws-sdk');
const XLSX = require('xlsx');
const fs = require('fs');
const cors = require('cors');
const sharp = require('sharp');

const app = express();
app.use(cors());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Cloudflare R2 S3-compatible config
const s3 = new AWS.S3({
endpoint: 'https://da20fb7146d20ac837f71d4abe3ae075.r2.cloudflarestorage.com',
accessKeyId: '76af20946b43fe93b87a439930413693',
secretAccessKey: '69c907799043353e70be6ff7721fefc41688b69bd53416febd9667bd58401643',
region: 'auto',
signatureVersion: 'v4',
});

const BUCKET_NAME = 'vivimart-admin-products';

app.post('/upload', upload.array('images'), async (req, res) => {
try {
const uploaded = [];

for (const file of req.files) {
  // Convert to webp buffer
  const webpBuffer = await sharp(file.buffer)
    .webp({ quality: 80 })  // You can adjust quality here (0-100)
    .toBuffer();

  // Change extension to .webp
  const filename = Date.now() + '-' + file.originalname.replace(/\.[^/.]+$/, "") + '.webp';

  await s3
    .putObject({
      Bucket: BUCKET_NAME,
      Key: filename,
      Body: webpBuffer,
      ContentType: 'image/webp',
    })
    .promise();

  const imageUrl = `https://pub-626045ddda6c426496bd466497a21725.r2.dev/${filename}`;
  uploaded.push({ name: file.originalname, url: imageUrl });
}

// Create Excel
const worksheet = XLSX.utils.json_to_sheet(uploaded);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Uploads');
const filePath = './image-urls.xlsx';
XLSX.writeFile(workbook, filePath);

// Send Excel file
res.download(filePath, 'image-urls.xlsx', err => {
  if (!err) fs.unlinkSync(filePath); // Clean up file
});
} catch (err) {
console.error('Upload failed:', err);
res.status(500).json({ error: 'Upload failed' });
}
});

app.listen(3000, '0.0.0.0', () => {
  console.log('Server running on port 3000');
});
