const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Banner = require("../model/banner.Schema");
const bannerController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");


bannerController.post("/create", upload.single("image"), async (req, res) => {
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
      obj = { ...req.body, image: image.url, type: req.body.type, };
    }
    const BannerCreated = await Banner.create(obj);
    sendResponse(res, 200, "Success", {
      message: "Banner created successfully!",
      data: BannerCreated,
      statusCode:200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode:500
    });
  }
});

bannerController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "", 
      status, 
      pageNo=1, 
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
    const bannerList = await Banner.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo-1) * parseInt(pageCount)); 
    const totalCount = await Banner.countDocuments({});
    const activeCount = await Banner.countDocuments({status:true});
    sendResponse(res, 200, "Success", {
      message: "Banner list retrieved successfully!",
      data: bannerList,
      documentCount: {totalCount, activeCount, inactiveCount: totalCount-activeCount},
      statusCode:200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode:500
    });
  }
});

bannerController.delete("/delete/:id", async (req, res) => {
    try {
      const { id } = req.params;
      const banner = await Banner.findById(id);
      if (!banner) {
        return sendResponse(res, 404, "Failed", {
          message: "Banner not found",
        });
      }
      const imageUrl = banner.image;
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
      await Banner.findByIdAndDelete(id);
      sendResponse(res, 200, "Success", {
        message: "Banner and associated image deleted successfully!",
        statusCode:200
      });
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
      });
    }
});

bannerController.put("/update", upload.single("image"), async (req, res) => {
  try {
    const  id  = req.body._id;
    const banner = await Banner.findById(id);
    if (!banner) {
      return sendResponse(res, 404, "Failed", {
        message: "Banner not found",
        statusCode:403
      });
    }
    let updatedData = { ...req.body };
    if (req.file) {
      // Delete the old image from Cloudinary
      if (banner.image) {
        const publicId = banner.image.split("/").pop().split(".")[0];
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
    const updatedBanner = await Banner.findByIdAndUpdate(id, updatedData, {
      new: true, // Return the updated document
    });
    sendResponse(res, 200, "Success", {
      message: "Banner updated successfully!",
      data: updatedBanner,
      statusCode:200
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode:500
    });
  }
});

module.exports = bannerController;
