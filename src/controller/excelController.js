const express = require("express");
require("dotenv").config();
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");
const mongoose = require("mongoose");
const { sendResponse } = require("../utils/common");
const upload = require("../utils/multer");
const cloudinary = require("../utils/cloudinary")

const Product = require("../model/product.Schema");
const Category = require("../model/category.Schema");
const Brand = require("../model/brand.Schema");

const excelController = express.Router();

// Convert string to ObjectId safely
const toObjectId = (id) => {
  try {
    return mongoose.Types.ObjectId(id.trim());
  } catch {
    return null;
  }
};

// Normalize each row before insert
const normalizeProductData = async (item) => {
  // CATEGORY: Convert category name to ID
  if (item.category && typeof item.category === "string") {
    const categoryNames = item.category.split(",").map(name => name.trim());
    const categoryDocs = await Category.find({ name: { $in: categoryNames } });
    item.categoryId = categoryDocs.map(cat => cat._id);
  }

  // BRAND: Convert brand name to ID
  if (item.brand && typeof item.brand === "string") {
    const brandDoc = await Brand.findOne({ name: item.brand.trim() });
    item.brandId = brandDoc ? brandDoc._id : null;
  }

  // PRODUCT HERO IMAGE: Upload from local path if provided
  if (item.productHeroImage && typeof item.productHeroImage === "string") {
    try {
      const uploadRes = await cloudinary.uploader.upload(item.productHeroImage, {
        folder: "products",
      });
      item.productHeroImage = uploadRes.secure_url;
    } catch (err) {
      console.error("Hero Image Upload Error:", err);
      item.productHeroImage = "";
    }
  } else {
    item.productHeroImage = "";
  }

  // PRODUCT GALLERY: Upload each local path
  if (item.productGallery && typeof item.productGallery === "string") {
    const galleryPaths = item.productGallery.split(",").map(url => url.trim());
    const uploadedGallery = [];

    for (const imgPath of galleryPaths) {
      try {
        const uploadRes = await cloudinary.uploader.upload(imgPath, {
          folder: "products/gallery",
        });
        uploadedGallery.push(uploadRes.secure_url);
      } catch (err) {
        console.error("Gallery Image Upload Error:", err);
      }
    }

    item.productGallery = uploadedGallery;
  } else {
    item.productGallery = [];
  }

  // TAGS: Parse comma separated tags
  if (item.tags && typeof item.tags === "string") {
    item.tags = item.tags.split(",").map(tag => tag.trim());
  } else {
    item.tags = [];
  }

  // SPECIAL APPEARANCE: Parse comma separated values
  if (item.specialAppearance && typeof item.specialAppearance === "string") {
    item.specialAppearance = item.specialAppearance.split(",").map(s => s.trim());
  } else {
    item.specialAppearance = [];
  }

  // Clean up unnecessary fields
  delete item.venderId;
  delete item.productOtherDetails;
  delete item.productVariants;
  delete item.category; // remove the category name field used for mapping
  delete item.brand;    // remove the brand name field used for mapping

  return item;
};

excelController.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return sendResponse(res, 400, "Failed", {
        message: "No file uploaded",
        statusCode: 400,
      });
    }

    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    fs.unlinkSync(filePath); // Clean up

    if (!jsonData || jsonData.length === 0) {
      return sendResponse(res, 422, "Failed", {
        message: "Excel file is empty or invalid",
        statusCode: 422,
      });
    }

    const processedData = await Promise.all(
      jsonData
        .filter(item => item.name && item.price)
        .map(async (item) => {
          // Check duplicate
          const existingProduct = await Product.findOne({ name: item.name.trim() });
          if (existingProduct) {
            const error = new Error(`Duplicate product name "${item.name}" found. Upload aborted.`);
            error.statusCode = 409; // Conflict
            throw error;
          }
    
          return normalizeProductData(item);
        })
    );
    
    

    // Insert into Product collection
    const insertedProducts = await Product.insertMany(processedData);

    return sendResponse(res, 200, "Success", {
      message: "Excel data uploaded and saved successfully!",
      data: insertedProducts,
      statusCode: 200,
    });

  } catch (error) {
    console.error("Excel Upload Error:", error);
    const statusCode = error.statusCode || 500;
    return sendResponse(res, statusCode, "Failed", {
      message: error.message || "Internal Server Error",
      statusCode,
    });
  }
  
});


excelController.post("/export", async (req, res) => {
  try {
    const { format = "excel" } = req.body;

    // Fetch all products with populated category and brand names if needed
    const products = await Product.find()
      .populate("categoryId", "name")
      .populate("brandId", "name")
      .lean();

      const processedProducts = products.map((p) => ({
        ...p,
        category: p.categoryId?.map(c => c.name).join(", ") || "",
        brand: p.brandId?.name || "",
        tags: Array.isArray(p.tags) ? p.tags.join(", ") : "",
        specialAppearance: Array.isArray(p.specialAppearance) ? p.specialAppearance.join(", ") : "",
      }));
      

    let fileBuffer;
    let contentType;
    let fileExtension;

    if (format === "excel") {
      // Convert to Excel file
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.json_to_sheet(processedProducts);
      xlsx.utils.book_append_sheet(workbook, worksheet, "Products");

      fileBuffer = xlsx.write(workbook, { type: "buffer", bookType: "xlsx" });
      contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      fileExtension = "xlsx";

    } else if (format === "csv") {
      // Convert to CSV
      const worksheet = xlsx.utils.json_to_sheet(processedProducts);
      fileBuffer = Buffer.from(xlsx.utils.sheet_to_csv(worksheet), "utf-8");
      contentType = "text/csv";
      fileExtension = "csv";

    } else if (format === "txt") {
      // Convert to TXT (tab separated)
      const worksheet = xlsx.utils.json_to_sheet(processedProducts);
      const txtData = xlsx.utils.sheet_to_txt(worksheet, { FS: "\t" });
      fileBuffer = Buffer.from(txtData, "utf-8");
      contentType = "text/plain";
      fileExtension = "txt";

    } else {
      return sendResponse(res, 400, "Failed", {
        message: "Invalid export format",
        statusCode: 400,
      });
    }

    // Set headers for download
    res.setHeader("Content-Type", contentType);
    res.setHeader("Content-Disposition", `attachment; filename=productList.${fileExtension}`);

    // Send the file buffer
    res.send(fileBuffer);

  } catch (error) {
    console.error("Export Error:", error);
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal Server Error",
      statusCode: 500,
    });
  }
});


module.exports = excelController;
