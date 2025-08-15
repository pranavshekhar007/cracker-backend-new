const express = require("express");
const { sendResponse } = require("../utils/common");
const ComboProduct = require("../model/comboProduct.Schema");
const comboProductController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const fs = require("fs");

// Create Combo Product
comboProductController.post("/create", async (req, res) => {
  try {
    const comboPrice = parseFloat(req.body?.pricing?.comboPrice || 0);

    if (comboPrice > 3000) {
      return sendResponse(res, 400, "Failed", {
        message: "Combo price cannot exceed ₹3000.",
        statusCode: 400,
      });
    }
    const comboProductData = {
      ...req.body,
    };

    const comboProductCreated = await ComboProduct.create(comboProductData);
    sendResponse(res, 200, "Success", {
      message: "Combo product created successfully!",
      data: comboProductCreated,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

// List Combo Products
comboProductController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      productType,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
      status,
    } = req.body;

    const query = {};
    if (productType) query.productType = productType;
    if (searchKey) query.name = { $regex: searchKey, $options: "i" };
    if (status) query.status = status;
    // Construct sorting object
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };

    // Fetch the combo product list
    const comboProductList = await ComboProduct.find(query)
      .populate("productId.product")
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await ComboProduct.countDocuments(query);
    const activeCount = await ComboProduct.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Combo product list retrieved successfully!",
      data: comboProductList,
      documentCount: {
        totalCount,
        activeCount,
        inactiveCount: totalCount - activeCount,
      },
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

// Update Combo Product
comboProductController.put("/update", async (req, res) => {
  try {
    const id = req.body.id;

    // Fetch the existing combo product by ID
    const comboProductData = await ComboProduct.findOne({ _id: id }).populate(
      "productId"
    );
    if (!comboProductData) {
      return sendResponse(res, 404, "Failed", {
        message: "Combo product not found",
      });
    }

    // Copy all fields from the request body
    let updateData = { ...req.body };

    // 🧼 Optional: Clean undefined/null/empty-string fields in update data
    Object.keys(updateData).forEach((key) => {
      if (
        updateData[key] === undefined ||
        updateData[key] === null ||
        (typeof updateData[key] === "string" && updateData[key].trim() === "")
      ) {
        delete updateData[key];
      }
    });

    // Ensure productId is always an array (even if a single product is provided)
    if (updateData.productId && !Array.isArray(updateData.productId)) {
      updateData.productId = [updateData.productId];
    }

    // ✅ Update combo product data in the database
    const updatedComboProduct = await ComboProduct.findByIdAndUpdate(
      id,
      updateData, // Use the cleaned update data
      { new: true } // Return the updated document
    );

    // If the update fails, return an error
    if (!updatedComboProduct) {
      return sendResponse(res, 404, "Failed", {
        message: "Combo product update failed",
        statusCode: 404,
      });
    }

    // Return a successful response with the updated combo product
    sendResponse(res, 200, "Success", {
      message: "Combo product updated successfully!",
      data: updatedComboProduct,
      statusCode: 200,
    });
  } catch (error) {
    console.error("Update error:", error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

// Delete Combo Product
comboProductController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;

    const deletedComboProduct = await ComboProduct.findByIdAndDelete(id);
    if (!deletedComboProduct) {
      return sendResponse(res, 404, "Failed", {
        message: "Combo product not found",
        statusCode: 404,
      });
    }
    sendResponse(res, 200, "Success", {
      message: "Combo product deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

comboProductController.get("/details/:id", async (req, res) => {
  try {
    const id = req.params.id;
    const product = await ComboProduct.findOne({ _id: id })
    .populate("productId.product");
    if (product) {
      return sendResponse(res, 200, "Success", {
        message: "Product details fetched  successfully",
        data: product,
        statusCode: 200,
      });
    } else {
      return sendResponse(res, 404, "Failed", {
        message: "Product not found",
        statusCode: 404,
      });
    }
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

comboProductController.put(
  "/update/hero-image",
  upload.single("productHeroImage"),
  async (req, res) => {
    try {
      const id = req.body.id;
      const product = await ComboProduct.findById(id);
      if (!product) {
        return sendResponse(res, 404, "Failed", {
          message: "Product not found",
          statusCode: 403,
        });
      }
      let updatedData = { ...req.body };
      if (req.file) {
        const productHeroImage = await cloudinary.uploader.upload(
          req.file.path
        );
        updatedData.productHeroImage = productHeroImage.url;
      }
      const updatedProduct = await ComboProduct.findByIdAndUpdate(
        id,
        updatedData,
        {
          new: true,
        }
      );
      sendResponse(res, 200, "Success", {
        message: "Product hero image updated successfully!",
        data: updatedProduct,
        statusCode: 200,
      });
    } catch (error) {
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500,
      });
    }
  }
);

comboProductController.put(
  "/update/add-product-gallery",
  upload.array("productGallery"),
  async (req, res) => {
    try {
      const id = req.body.id;

      const product = await ComboProduct.findById(id);
      if (!product) {
        return sendResponse(res, 404, "Failed", {
          message: "Product not found",
          statusCode: 403,
        });
      }

      if (!req.files || req.files.length === 0) {
        return sendResponse(res, 400, "Failed", {
          message: "At least one image file is required",
          statusCode: 400,
        });
      }

      // Upload all images to Cloudinary
      const uploadedUrls = [];
      for (const file of req.files) {
        const uploadedImage = await cloudinary.uploader.upload(file.path);
        uploadedUrls.push(uploadedImage.secure_url);
      }

      // Push all URLs into productGallery
      const updatedProduct = await ComboProduct.findByIdAndUpdate(
        id,
        { $push: { productGallery: { $each: uploadedUrls } } },
        { new: true }
      );

      sendResponse(res, 200, "Success", {
        message: "Product gallery images added successfully!",
        data: updatedProduct,
        statusCode: 200,
      });
    } catch (error) {
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500,
      });
    }
  }
);
comboProductController.post("/delete/product-gallery", async (req, res) => {
  try {
    const { id, index } = req.body;

    if (!id || index === undefined) {
      return sendResponse(res, 400, "Failed", {
        message: "Product ID and index are required",
        statusCode: 400,
      });
    }

    const product = await ComboProduct.findById(id);
    if (!product) {
      return sendResponse(res, 404, "Failed", {
        message: "Product not found",
        statusCode: 404,
      });
    }

    const gallery = product.productGallery;

    if (!gallery || index < 0 || index >= gallery.length) {
      return sendResponse(res, 400, "Failed", {
        message: "Invalid image index",
        statusCode: 400,
      });
    }

    const imageUrl = gallery[index];

    // Delete from Cloudinary if image is stored there
    const publicId = imageUrl.split("/").pop().split(".")[0];
    await cloudinary.uploader.destroy(publicId, (error, result) => {
      if (error) {
        console.error("Cloudinary delete error:", error);
      } else {
        console.log("Image deleted from Cloudinary:", result);
      }
    });

    // Remove the image from productGallery
    gallery.splice(index, 1);
    product.productGallery = gallery;
    const updatedProduct = await product.save();

    sendResponse(res, 200, "Success", {
      message: "Gallery image deleted successfully",
      data: updatedProduct,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal Server Error",
      statusCode: 500,
    });
  }
});

comboProductController.put(
  "/update-video",
  upload.single("productVideo"),
  async (req, res) => {
    try {
      const { id } = req.body;

      // Check if product exists
      const product = await ComboProduct.findById(id);
      if (!product) {
        return sendResponse(res, 404, "Failed", {
          message: "Product not found",
          statusCode: 404,
        });
      }

      // Initialize updatedData
      let updatedData = {};

      // Check if video file is uploaded
      if (req.file) {
        // Validate file type
        if (!req.file.mimetype.startsWith("video/")) {
          // Clean up the uploaded file
          fs.unlinkSync(req.file.path);
          return sendResponse(res, 400, "Failed", {
            message: "Only video files are allowed",
            statusCode: 400,
          });
        }

        // Upload to Cloudinary with resource_type set to "video"
        const productVideo = await cloudinary.uploader.upload(req.file.path, {
          resource_type: "video",
        });

        // Set video URL in update
        updatedData.productVideo = productVideo.url;

        // Delete local file after upload
        fs.unlinkSync(req.file.path);
      }

      // Update the product
      const updatedProduct = await ComboProduct.findByIdAndUpdate(
        id,
        updatedData,
        {
          new: true,
        }
      );

      // Send success response
      sendResponse(res, 200, "Success", {
        message: "Product video updated successfully!",
        data: updatedProduct,
        statusCode: 200,
      });
    } catch (error) {
      console.error("Update video error:", error.message);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500,
      });
    }
  }
);

module.exports = comboProductController;
