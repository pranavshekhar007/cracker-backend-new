const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Rating = require("../model/productRating.Schema");
const Product = require("../model/product.Schema");
const productRatingController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const auth = require("../utils/auth");


productRatingController.post("/create", async (req, res) => {
  try {
    

    const ratingCreated = await Rating.create(req.body);

    // ⭐ Calculate the average rating of this product
    const productId = ratingCreated.productId;

    const allRatings = await Rating.find({ productId });

    const totalRatings = allRatings.length;
    const sumRatings = allRatings.reduce((sum, item) => sum + Number(item.rating), 0);

    const averageRating = (sumRatings / totalRatings).toFixed(1); // 1 decimal point ka average

    // ⭐ Update product's rating field
    await Product.findByIdAndUpdate(productId, {
      rating: averageRating,
    });

    sendResponse(res, 200, "Success", {
      message: "Your review has been submitted successfully!",
      data: ratingCreated,
      averageRating,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});
productRatingController.post("/list", async (req, res) => {
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

    // Construct sorting object
    const sortField = sortByField || "createdAt"; 
    const sortOrder = sortByOrder === "asc" ? 1 : -1; 
    const sortOption = { [sortField]: sortOrder };

    // Fetch the rating list
    const ratingList = await Rating.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo-1) * parseInt(pageCount))
      .populate({
        path: "productId", // Field to populate
        select: "name description", // Specify the fields to retrieve from the rating collection
      })
      .populate({
        path: "userId", // Field to populate
         // Specify the fields to retrieve from the rating collection
      })
    const totalCount = await Rating.countDocuments({});
    const activeCount = await Rating.countDocuments({status:true});
    sendResponse(res, 200, "Success", {
      message: "Rating list retrieved successfully!",
      data: ratingList,
      documentCount: {totalCount, activeCount, inactiveCount: totalCount-activeCount},
      statusCode:200

    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

productRatingController.put("/update", upload.single("image"), async (req, res) => {
  try {
    const id = req.body._id;

    // Find the rating by ID
    const ratingData = await Rating.findById(id);
    if (!ratingData) {
      return sendResponse(res, 404, "Failed", {
        message: "rating not found",
      });
    }

    let updatedData = { ...req.body };

    // If a new image is uploaded
    if (req.file) {
      // Delete the old image from Cloudinary
      if (ratingData.image) {
        const publicId = ratingData.image.split("/").pop().split(".")[0];
        await cloudinary.uploader.destroy(publicId, (error, result) => {
          if (error) {
            console.error("Error deleting old image from Cloudinary:", error);
          } else {
            console.log("Old image deleted from Cloudinary:", result);
          }
        });
      }

      // Upload the new image to Cloudinary
      const image = await cloudinary.uploader.upload(req.file.path);
      updatedData.image = image.url;
    }

    // Update the rating in the database
    const updatedRating = await Rating.findByIdAndUpdate(id, updatedData, {
      new: true, // Return the updated document
    });

    sendResponse(res, 200, "Success", {
      message: "Rating updated successfully!",
      data: updatedRating,
      statusCode:200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

productRatingController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    console.log(id);
    // Find the rating by ID
    const ratingItem = await Rating.findById(id);
    if (!ratingItem) {
      return sendResponse(res, 404, "Failed", {
        message: "Rating not found",
      });
    }

    // Extract the public ID from the Cloudinary image URL
    const imageUrl = Rating.image;
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

    // Delete the rating from the database
    await Rating.findByIdAndDelete(id);

    sendResponse(res, 200, "Success", {
      message: "Rating and associated image deleted successfully!",
      statusCode:200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

module.exports = productRatingController;