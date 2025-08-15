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

const isHttpUrl = (s) => typeof s === "string" && /^https?:\/\//i.test(s);
const isDataUri = (s) => typeof s === "string" && s.startsWith("data:");

// const uploadLocalImage = async (filePath, folder) => {
//   console.log("file path: ", filePath);
//   console.log("Folder", folder);
//   try {
//     // Ensure path is valid for server
//     if (!fs.existsSync(filePath)) {
//       console.error("File not found:", filePath);
//       return "";
//     }

//     // Read file as base64
//     const fileData = fs.readFileSync(filePath, { encoding: "base64" });

//     console.log("File Data", fileData);

//     // Guess extension (default jpeg if unknown)
//     const ext = path.extname(filePath).replace(".", "") || "jpeg";

//     console.log("extension:", ext);

//     // Convert to Data URI
//     const dataURI = `data:image/${ext};base64,${fileData}`;

//     console.log("data", dataURI);

//     // Upload to Cloudinary
//     const uploadRes = await cloudinary.uploader.upload(dataURI, { folder });
//     console.log("upload",uploadRes);
//     return uploadRes.secure_url;
//   } catch (err) {
//     console.error(`Image upload error for ${filePath}:`, err);
//     return "";
//   }
// };

// Convert string to ObjectId safely
// const toObjectId = (id) => {
//   try {
//     return mongoose.Types.ObjectId(id.trim());
//   } catch {
//     return null;
//   }
// };

// Normalize each row before insert


const uploadImage = async (input, folder) => {
  try {
    if (!input || typeof input !== "string") return "";

    if (isHttpUrl(input) || isDataUri(input)) {
      const res = await cloudinary.uploader.upload(input, { folder });
      return res.secure_url;
    }

    // Optional: allow relative paths that exist on the server bundle (rare)
    // CAUTION: In serverless environments, only files bundled or in /tmp exist.
    // If you want to support relative paths during local dev only:
    if (!path.isAbsolute(input)) {
      const abs = path.join(process.cwd(), input);
      if (fs.existsSync(abs)) {
        const ext = path.extname(abs).replace(".", "") || "jpeg";
        const fileData = fs.readFileSync(abs, { encoding: "base64" });
        const dataURI = `data:image/${ext};base64,${fileData}`;
        const res = await cloudinary.uploader.upload(dataURI, { folder });
        return res.secure_url;
      }
    }

    console.error("Unsupported path in serverless (use URL or data URI):", input);
    return "";
  } catch (err) {
    console.error("Cloudinary upload error:", err);
    return "";
  }
};

// Convert string to ObjectId safely
const toObjectId = (id) => {
  try {
    return mongoose.Types.ObjectId(id.trim());
  } catch {
    return null;
  }
};

// Normalize each row before insert/update
const normalizeProductData = async (item) => {
  // CATEGORY: Convert category name to IDs
  if (item.category && typeof item.category === "string") {
    const categoryNames = item.category.split(",").map((name) => name.trim());
    const categoryDocs = await Category.find({ name: { $in: categoryNames } });
    item.categoryId = categoryDocs.map((cat) => cat._id);
  }

  // BRAND: Convert brand name to ID
  if (item.brand && typeof item.brand === "string") {
    const brandDoc = await Brand.findOne({ name: item.brand.trim() });
    item.brandId = brandDoc ? brandDoc._id : null;
  }

  // PRODUCT HERO IMAGE
  if (item.productHeroImage && typeof item.productHeroImage === "string") {
    item.productHeroImage = await uploadImage(item.productHeroImage, "products");
  } else {
    item.productHeroImage = "";
  }

  // PRODUCT GALLERY
  if (item.productGallery && typeof item.productGallery === "string") {
    const galleryInputs = item.productGallery
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);

    const uploadedGallery = [];
    for (const src of galleryInputs) {
      try {
        const url = await uploadImage(src, "products/gallery");
        if (url) uploadedGallery.push(url);
      } catch (err) {
        console.error("Gallery Image Upload Error:", err);
      }
    }
    item.productGallery = uploadedGallery;
  } else {
    item.productGallery = [];
  }

  // TAGS
  if (item.tags && typeof item.tags === "string") {
    item.tags = item.tags.split(",").map((tag) => tag.trim());
  } else {
    item.tags = [];
  }

  // SPECIAL APPEARANCE
  if (item.specialAppearance && typeof item.specialAppearance === "string") {
    item.specialAppearance = item.specialAppearance
      .split(",")
      .map((s) => s.trim());
  } else {
    item.specialAppearance = [];
  }

  // Clean up unnecessary fields
  delete item.venderId;
  delete item.productOtherDetails;
  delete item.productVariants;
  delete item.category; // remove the category name field used for mapping
  delete item.brand; // remove the brand name field used for mapping

  return item;
};

excelController.post("/upload-or-update", upload.single("file"), async (req, res) => {
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

    // Clean up temp file (multer disk storage)
    try {
      fs.unlinkSync(filePath);
    } catch (e) {
      console.warn("Failed to remove temp file:", e?.message);
    }

    if (!jsonData || jsonData.length === 0) {
      return sendResponse(res, 422, "Failed", {
        message: "Excel file is empty or invalid",
        statusCode: 422,
      });
    }

    const insertedProducts = [];
    const updatedProducts = [];

    for (const item of jsonData.filter((i) => i.name && i.price)) {
      const normalizedItem = await normalizeProductData({ ...item });

      const existingProduct = await Product.findOne({ name: item.name.trim() });
      if (existingProduct) {
        await Product.updateOne({ _id: existingProduct._id }, { $set: normalizedItem });
        updatedProducts.push(existingProduct.name);
      } else {
        const newProduct = await Product.create(normalizedItem);
        insertedProducts.push(newProduct.name);
      }
    }

    let message = "";
    if (insertedProducts.length > 0 && updatedProducts.length > 0) {
      message = "Products uploaded and updated successfully!";
    } else if (insertedProducts.length > 0) {
      message = "Products uploaded successfully!";
    } else if (updatedProducts.length > 0) {
      message = "Products updated successfully!";
    } else {
      message = "No products were uploaded or updated.";
    }

    return sendResponse(res, 200, "Success", {
      message,
      insertedCount: insertedProducts.length,
      inserted: insertedProducts,
      updatedCount: updatedProducts.length,
      updated: updatedProducts,
      statusCode: 200,
    });
  } catch (error) {
    console.error("Excel Upload/Update Error:", error);
    const statusCode = error.statusCode || 500;
    return sendResponse(res, statusCode, "Failed", {
      message: error.message || "Internal Server Error",
      statusCode,
    });
  }
});

// const normalizeProductData = async (item) => {
//   // CATEGORY: Convert category name to ID
//   if (item.category && typeof item.category === "string") {
//     const categoryNames = item.category.split(",").map(name => name.trim());
//     const categoryDocs = await Category.find({ name: { $in: categoryNames } });
//     item.categoryId = categoryDocs.map(cat => cat._id);
//   }

//   // BRAND: Convert brand name to ID
//   if (item.brand && typeof item.brand === "string") {
//     const brandDoc = await Brand.findOne({ name: item.brand.trim() });
//     item.brandId = brandDoc ? brandDoc._id : null;
//   }

//   // PRODUCT HERO IMAGE: Upload from local path if provided
//   // if (item.productHeroImage && typeof item.productHeroImage === "string") {
//   //   try {
//   //     const uploadRes = await cloudinary.uploader.upload(item.productHeroImage, {
//   //       folder: "products",
//   //     });
//   //     console.log("uploade url: ", uploadRes.secure_url)
//   //     item.productHeroImage = uploadRes.secure_url;
//   //   } catch (err) {
//   //     console.error("Hero Image Upload Error:", err);
//   //     item.productHeroImage = "";
//   //   }
//   // } else {
//   //   item.productHeroImage = "";
//   // }

//   if (item.productHeroImage && typeof item.productHeroImage === "string") {
//     item.productHeroImage = await uploadLocalImage(item.productHeroImage, "products");
//   } else {
//     item.productHeroImage = "";
//   }

//   // PRODUCT GALLERY: Upload each local path
//   if (item.productGallery && typeof item.productGallery === "string") {
//     const galleryPaths = item.productGallery.split(",").map(url => url.trim());
//     const uploadedGallery = [];

//     for (const imgPath of galleryPaths) {
//       try {
//         const uploadRes = await cloudinary.uploader.upload(imgPath, {
//           folder: "products/gallery",
//         });
//         uploadedGallery.push(uploadRes.secure_url);
//       } catch (err) {
//         console.error("Gallery Image Upload Error:", err);
//       }
//     }

//     item.productGallery = uploadedGallery;
//   } else {
//     item.productGallery = [];
//   }

//   // TAGS: Parse comma separated tags
//   if (item.tags && typeof item.tags === "string") {
//     item.tags = item.tags.split(",").map(tag => tag.trim());
//   } else {
//     item.tags = [];
//   }

//   // SPECIAL APPEARANCE: Parse comma separated values
//   if (item.specialAppearance && typeof item.specialAppearance === "string") {
//     item.specialAppearance = item.specialAppearance.split(",").map(s => s.trim());
//   } else {
//     item.specialAppearance = [];
//   }

//   // Clean up unnecessary fields
//   delete item.venderId;
//   delete item.productOtherDetails;
//   delete item.productVariants;
//   delete item.category; // remove the category name field used for mapping
//   delete item.brand;    // remove the brand name field used for mapping

//   return item;
// };

// excelController.post("/upload-or-update", upload.single("file"), async (req, res) => {
//   try {
//     if (!req.file) {
//       return sendResponse(res, 400, "Failed", {
//         message: "No file uploaded",
//         statusCode: 400,
//       });
//     }

//     const filePath = req.file.path;
//     const workbook = xlsx.readFile(filePath);
//     const sheetName = workbook.SheetNames[0];
//     const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

//     fs.unlinkSync(filePath); // Clean up temp file

//     if (!jsonData || jsonData.length === 0) {
//       return sendResponse(res, 422, "Failed", {
//         message: "Excel file is empty or invalid",
//         statusCode: 422,
//       });
//     }

//     let insertedProducts = [];
//     let updatedProducts = [];

//     for (const item of jsonData.filter(i => i.name && i.price)) {
//       // Normalize data
//       const normalizedItem = await normalizeProductData({ ...item });

//       // Check if product exists by name
//       const existingProduct = await Product.findOne({ name: item.name.trim() });

//       if (existingProduct) {
//         // Update product
//         await Product.updateOne(
//           { _id: existingProduct._id },
//           { $set: normalizedItem }
//         );
//         updatedProducts.push(existingProduct.name);
//       } else {
//         // Insert new product
//         const newProduct = await Product.create(normalizedItem);
//         insertedProducts.push(newProduct.name);
//       }
//     }

//     // Dynamic message
//     let message = "";
//     if (insertedProducts.length > 0 && updatedProducts.length > 0) {
//       message = "Products uploaded and updated successfully!";
//     } else if (insertedProducts.length > 0) {
//       message = "Products uploaded successfully!";
//     } else if (updatedProducts.length > 0) {
//       message = "Products updated successfully!";
//     } else {
//       message = "No products were uploaded or updated.";
//     }

//     return sendResponse(res, 200, "Success", {
//       message,
//       insertedCount: insertedProducts.length,
//       inserted: insertedProducts,
//       updatedCount: updatedProducts.length,
//       updated: updatedProducts,
//       statusCode: 200,
//     });

//   } catch (error) {
//     console.error("Excel Upload/Update Error:", error);
//     const statusCode = error.statusCode || 500;
//     return sendResponse(res, statusCode, "Failed", {
//       message: error.message || "Internal Server Error",
//       statusCode,
//     });
//   }
// });




excelController.post("/export", async (req, res) => {
  try {
    const { format = "excel" } = req.body;

    // Fetch products & populate category and brand
    const products = await Product.find()
      .populate("categoryId", "name")
      .populate("brandId", "name")
      .lean();

    // Map to match upload template format
    const processedProducts = products.map((p) => ({
      name: p.name || "",
      tags: Array.isArray(p.tags) ? p.tags.join(", ") : "",
      category: Array.isArray(p.categoryId) ? p.categoryId.map(c => c.name).join(", ") : "",
      brand: p.brandId?.name || "",
      specialAppearance: Array.isArray(p.specialAppearance) ? p.specialAppearance.join(", ") : "",
      shortDescription: p.shortDescription || "",
      stockQuantity: p.stockQuantity || 0,
      price: p.price || 0,
      discountedPrice: p.discountedPrice || 0,
      numberOfPieces: p.numberOfPieces || "",
      soundLevel: p.soundLevel || "",
      lightEffect: p.lightEffect || "",
      safetyRating: p.safetyRating || "",
      usageArea: p.usageArea || "",
      duration: p.duration || "",
      weightPerBox: p.weightPerBox || "",
      productHeroImage: p.productHeroImage || "",
      productGallery: Array.isArray(p.productGallery) ? p.productGallery.join(", ") : "",
      status: p.status ? "True" : "False",
    }));

    let fileBuffer;
    let contentType;
    let fileExtension;

    if (format === "excel") {
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.json_to_sheet(processedProducts, { header: Object.keys(processedProducts[0] || {}) });
      xlsx.utils.book_append_sheet(workbook, worksheet, "Products");
      fileBuffer = xlsx.write(workbook, { type: "buffer", bookType: "xlsx" });
      contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      fileExtension = "xlsx";
    } else if (format === "csv") {
      const worksheet = xlsx.utils.json_to_sheet(processedProducts);
      fileBuffer = Buffer.from(xlsx.utils.sheet_to_csv(worksheet), "utf-8");
      contentType = "text/csv";
      fileExtension = "csv";
    } else if (format === "txt") {
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

    // Send file
    res.setHeader("Content-Type", contentType);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=BulkProductUploadTemplate.${fileExtension}`
    );
    res.send(fileBuffer);

  } catch (error) {
    console.error("Export Error:", error);
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal Server Error",
      statusCode: 500,
    });
  }
});


excelController.get("/sample", async (req, res) => {
  try {
    const format = (req.query.format || "excel").toLowerCase();

    // Headers as per BulkProductUploadTemplate.xlsx
    const headers = [
      "name",
      "tags",
      "category",
      "brand",
      "specialAppearance",
      "shortDescription",
      "stockQuantity",
      "price",
      "discountedPrice",
      "numberOfPieces",
      "soundLevel",
      "lightEffect",
      "safetyRating",
      "usageArea",
      "duration",
      "weightPerBox",
      "productHeroImage",
      "productGallery",
      "status"
    ];

    // Create a blank worksheet with just headers
    const ws = xlsx.utils.aoa_to_sheet([headers]);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "Sample");

    let fileBuffer;
    let contentType;
    let fileExtension;

    if (format === "excel") {
      fileBuffer = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
      contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      fileExtension = "xlsx";
    } 
    else if (format === "csv") {
      fileBuffer = Buffer.from(xlsx.utils.sheet_to_csv(ws), "utf-8");
      contentType = "text/csv";
      fileExtension = "csv";
    } 
    else if (format === "txt") {
      const txtData = xlsx.utils.sheet_to_txt(ws, { FS: "\t" });
      fileBuffer = Buffer.from(txtData, "utf-8");
      contentType = "text/plain";
      fileExtension = "txt";
    } 
    else {
      return sendResponse(res, 400, "Failed", {
        message: "Invalid format. Use excel, csv, or txt.",
        statusCode: 400,
      });
    }

    res.setHeader("Content-Type", contentType);
    res.setHeader("Content-Disposition", `attachment; filename=BulkProductUploadTemplate.${fileExtension}`);
    res.send(fileBuffer);

  } catch (error) {
    console.error("Sample File Download Error:", error);
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal Server Error",
      statusCode: 500,
    });
  }
});


module.exports = excelController;
