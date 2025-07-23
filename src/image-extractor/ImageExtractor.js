import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import './ImageExtractor.css';

const ImageExtractor = () => {
  const [images, setImages] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const handleImageChange = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    setIsLoading(true);

    try {
      const newImages = await Promise.all(
        files.map(async (file) => {
          const url = URL.createObjectURL(file);
          const dimensions = await getImageSize(url);
          return {
            name: file.name,
            url,
            file,
            width: dimensions.width,
            height: dimensions.height,
          };
        })
      );

      setImages([...images, ...newImages]);
    } catch (error) {
      console.error('Error processing images:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const exportToExcel = async () => {
    if (images.length === 0) return;

    setIsExporting(true);

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Images');

      worksheet.columns = [
        { header: 'S.No', key: 'id', width: 8 },
        { header: 'Image Name', key: 'name', width: 40 },
        { header: 'Preview', key: 'image', width: 30 },
      ];

      // ✅ Header styling
      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4CAF50' },
      };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
      headerRow.height = 25;

      for (let i = 0; i < images.length; i++) {
        const img = images[i];
        const rowNumber = i + 2;

        worksheet.addRow({
          id: i + 1,
          name: img.name,
        });

        // ✅ Make cell text bigger & wrap
        worksheet.getCell(`A${rowNumber}`).font = { size: 12 };
        worksheet.getCell(`B${rowNumber}`).font = { size: 12 };
        worksheet.getCell(`B${rowNumber}`).alignment = { wrapText: true }; // wrap image name

        // ✅ Adjust image cell size
        const maxCellWidthPx = 180; // bigger
        const maxCellHeightPx = 150; // bigger row height

        const scale = Math.min(
          maxCellWidthPx / img.width,
          maxCellHeightPx / img.height
        );

        const displayWidth = img.width * scale;
        const displayHeight = img.height * scale;

        const row = worksheet.getRow(rowNumber);
        row.height = displayHeight / 1.33;

        const base64Image = await getBase64Image(img.file);

        const imageId = workbook.addImage({
          base64: base64Image,
          extension: img.name.split('.').pop(),
        });

        worksheet.addImage(imageId, {
          tl: { col: 2, row: rowNumber - 1, offsetX: 5, offsetY: 5 },
          ext: { width: displayWidth, height: displayHeight },
        });

        worksheet.getCell(`A${rowNumber}`).alignment = { horizontal: 'center' };
      }

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), 'image_list.xlsx');
    } catch (error) {
      console.error('Error exporting to Excel:', error);
    } finally {
      setIsExporting(false);
    }
  };

  const getImageSize = (url) => {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        resolve({ width: img.width, height: img.height });
      };
      img.src = url;
    });
  };

  const getBase64Image = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        resolve(e.target.result.split(',')[1]);
      };
      reader.readAsDataURL(file);
    });
  };

  const removeImage = (index) => {
    const newImages = [...images];
    URL.revokeObjectURL(newImages[index].url);
    newImages.splice(index, 1);
    setImages(newImages);
  };

  const removeAllImages = () => {
    images.forEach((img) => URL.revokeObjectURL(img.url));
    setImages([]);
  };

  return (
    <div className="image-extractor-container">
      <h1 className="image-extractor-header">Image Extractor to Excel</h1>

      <div className="button-container">
        <input
          type="file"
          accept="image/*"
          multiple
          onChange={handleImageChange}
          className="file-input"
          id="imageInput"
          disabled={isLoading || isExporting}
        />
        <label
          htmlFor="imageInput"
          className={`primary-button ${isLoading || isExporting ? 'disabled' : ''}`}
        >
          {isLoading ? 'Uploading...' : 'Select Images'}
        </label>

        <button
          onClick={exportToExcel}
          disabled={images.length === 0 || isExporting || isLoading}
          className={`primary-button export-button ${images.length ? '' : 'disabled'}`}
        >
          {isExporting ? 'Exporting...' : 'Export to Excel'}
        </button>

        {images.length > 0 && (
          <button
            onClick={removeAllImages}
            className="secondary-button"
            disabled={isLoading || isExporting}
          >
            Clear All
          </button>
        )}
      </div>

      {isLoading ? (
        <div className="loading-state">
          <div className="spinner"></div>
          <p>Processing images...</p>
        </div>
      ) : images.length === 0 ? (
        <div className="empty-state">
          <div className="empty-state-content">
            <p className="empty-state-text">No images selected</p>
            <p className="empty-state-subtext">Click "Select Images" to add images</p>
            <label htmlFor="imageInput" className="empty-state-button">
              Select Images
            </label>
          </div>
        </div>
      ) : (
        <div className="preview-container">
          {images.map((img, index) => (
            <div key={index} className="image-card">
              <img src={img.url} alt={img.name} className="preview-image" />
              <p className="file-name">{img.name}</p>
              <button
                onClick={() => removeImage(index)}
                className="remove-button"
                title="Remove image"
              >
                ×
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default ImageExtractor;
