const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const subCategory = require("../model/subCategory.Schema");
const Product = require("../model/product.Schema");
const AttributeSet = require("../model/attributeSet.Schema");
const Attribute = require("../model/attribute.Schema");
const subCategoryController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const auth = require("../utils/auth");

subCategoryController.post(
  "/create",
  upload.single("image"),
  async (req, res) => {
    try {
      let obj;
      if (req.file) {
        let image = await cloudinary.uploader.upload(
          req.file.path,
          function (err, result) {
            if (err) {
              return err;
            } else {
              return result;
            }
          }
        );
        obj = { ...req.body, image: image.url };
      }
      const subCategoryCreated = await subCategory.create(obj);
      sendResponse(res, 200, "Success", {
        message: "Sub Category created successfully!",
        data: subCategoryCreated,
        statusCode: 200,
      });
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
      });
    }
  }
);

subCategoryController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      status,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;

    const query = {};
    if (status) query.status = status;
    if (searchKey) query.name = { $regex: searchKey, $options: "i" };

    // Construct sorting object
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };

    // Fetch the category list
    const subCategoryList = await subCategory
      .find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount))
      .populate({
        path: "categoryId", // Field to populate
        select: "name description", // Specify the fields to retrieve from the category collection
      });
    const totalCount = await subCategory.countDocuments({});
    const activeCount = await subCategory.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Sub Category list retrieved successfully!",
      data: subCategoryList,
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

subCategoryController.post("/attribute-list", async (req, res) => {
  try {
    const { productId } = req.body;

    const productDetails = await Product.findOne({ _id: productId });

    if (!productDetails) {
      return sendResponse(res, 404, "Failed", {
        message: "Product not found",
      });
    }

    // get list of attributeSets under this product's subcategory
    const attributeSetList = await AttributeSet.find({ subCategoryId: productDetails.subCategoryId });
    const attributeSetIds = attributeSetList.map((set) => set._id);

    // get all attributes under those sets
    const attributeList = await Attribute.find({
      attributeSetId: { $in: attributeSetIds },
    });

    // extract keys already added in productOtherDetails
    const alreadyAddedAttributeNames = productDetails.productOtherDetails.map((detail) => detail.key);
    // filter out already added attributes
    const filteredAttributes = attributeList.filter(attr => !alreadyAddedAttributeNames.includes(attr?.name));

    // format result
    const formattedAttributes = filteredAttributes.map((attr) => ({
      attributeId: attr._id,
      name: attr.name,
      value: attr.value,
      attributeSetId: attr.attributeSetId,
      status: attr.status,
    }));

    // send response
    sendResponse(res, 200, "Success", {
      message: "Attribute list retrieved successfully!",
      data: formattedAttributes,
      statusCode: 200,
    });

  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

subCategoryController.put(
  "/update",
  upload.single("image"),
  async (req, res) => {
    try {
      const id = req.body._id;

      // Find the category by ID
      const subCategoryData = await subCategory.findById(id);
      if (!subCategoryData) {
        return sendResponse(res, 404, "Failed", {
          message: "Sub Category not found",
        });
      }

      let updatedData = { ...req.body };

      // If a new image is uploaded
      if (req.file) {
        // Delete the old image from Cloudinary
        if (subCategoryData.image) {
          const publicId = subCategory.image.split("/").pop().split(".")[0];
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

      // Update the category in the database
      const updatedSubCategory = await subCategory.findByIdAndUpdate(
        id,
        updatedData,
        {
          new: true, // Return the updated document
        }
      );

      sendResponse(res, 200, "Success", {
        message: "Sub Category updated successfully!",
        data: updatedSubCategory,
        statusCode: 200,
      });
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
      });
    }
  }
);

subCategoryController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    console.log(id);
    // Find the category by ID
    const subCategoryItem = await subCategory.findById(id);
    if (!subCategoryItem) {
      return sendResponse(res, 404, "Failed", {
        message: "Sub Category not found",
      });
    }

    // Extract the public ID from the Cloudinary image URL
    const imageUrl = subCategory.image;
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

    // Delete the category from the database
    await subCategory.findByIdAndDelete(id);

    sendResponse(res, 200, "Success", {
      message: "Sub Category and associated image deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

subCategoryController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const subCategoryDetails = await subCategory.findOne({ _id: id });
    const productList = await Product.find({ subCategoryId: id });
    sendResponse(res, 200, "Success", {
      message: "Sub category with product list retrived successfully!",
      data: { subCategoryDetails, productList },
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

module.exports = subCategoryController;
