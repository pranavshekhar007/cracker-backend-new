const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Brand = require("../model/brand.Schema");
const brandController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");



brandController.post("/create", upload.single("image"), async (req, res) => {
  try {
    let obj;
    if (req.file) {
      let image = await cloudinary.uploader.upload(req.file.path, function (err, result) {
        if (err) {
          return err;
        } else {
          return result;
        }
      });
      obj = { ...req.body, image: image.url };
    }
    const BrandCreated = await Brand.create(obj);
    sendResponse(res, 200, "Success", {
      message: "Brand created successfully!",
      data: BrandCreated,
      statusCode: 200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500
    });
  }
});


brandController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      status,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder
    } = req.body;
    const query = {};
    if (status) query.status = status;
    if (searchKey) query.name = { $regex: searchKey, $options: "i" };
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };
    const brandList = await Brand.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await Brand.countDocuments({});
    const activeCount = await Brand.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Brand list retrieved successfully!",
      data: brandList,
      documentCount: { totalCount, activeCount, inactiveCount: totalCount - activeCount },
      statusCode: 200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500
    });
  }
});


brandController.put("/update", upload.single("image"), async (req, res) => {
  try {
    const id = req.body._id;
    const brand = await Brand.findById(id);
    if (!brand) {
      return sendResponse(res, 404, "Failed", {
        message: "Brand not found",
        statusCode: 403
      });
    }
    let updatedData = { ...req.body };
    if (req.file) {
      // Delete the old image from Cloudinary
      if (brand.image) {
        const publicId = brand.image.split("/").pop().split(".")[0];
        await cloudinary.uploader.destroy(publicId, (error, result) => {
          if (error) {
            console.error("Error deleting old image from Cloudinary:", error);
          } else {
            console.log("Old image deleted from Cloudinary:", result);
          }
        });
      }
      const image = await cloudinary.uploader.upload(req.file.path);
      updatedData.image = image.url;
    }
    const updatedBrand = await Brand.findByIdAndUpdate(id, updatedData, {
      new: true, // Return the updated document
    });
    sendResponse(res, 200, "Success", {
      message: "Brand updated successfully!",
      data: updatedBrand,
      statusCode: 200
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500
    });
  }
});


brandController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const brand = await Brand.findById(id);
    if (!brand) {
      return sendResponse(res, 404, "Failed", {
        message: "Brand not found",
      });
    }
    const imageUrl = brand.image;
    if (imageUrl) {
      const publicId = imageUrl.split("/").pop().split(".")[0]; // Extract public ID
      // Delete the image from Cloudinary
      await cloudinary.uploader.destroy(publicId, (error, result) => {
        if (error) {
          console.error("Error deleting image from Cloudinary:", error);
        } else {
          console.log("Cloudinary image deletion result:", result);
        }
      });
    }
    await Brand.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Brand associated image deleted successfully!",
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});


brandController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params
    const brandDetails = await Brand.findOne({ _id: id });
    sendResponse(res, 200, "Success", {
      message: "Brand retrived successfully!",
      data: { brandDetails },
      statusCode: 200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500
    });
  }
});


module.exports = brandController;