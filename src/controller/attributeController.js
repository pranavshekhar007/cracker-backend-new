const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Attribute = require("../model/attribute.Schema");
const AttributeSet = require("../model/attributeSet.Schema");
const attributeController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const auth = require("../utils/auth");

attributeController.post("/create", async (req, res) => {
  try {
    const attributeCreated = await Attribute.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Attribute created successfully!",
      data: attributeCreated,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

attributeController.post("/list", async (req, res) => {
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
    const attributeList = await Attribute.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount))
      .populate({
        path: "attributeSetId", // Field to populate
        select: "name description", // Specify the fields to retrieve from the category collection
      });
    const totalCount = await Attribute.countDocuments({});
    const activeCount = await Attribute.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Attribute list retrieved successfully!",
      data: attributeList,
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

attributeController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;

    // Find the category by ID
    const attributeData = await Attribute.findById(id);
    if (!attributeData) {
      return sendResponse(res, 404, "Failed", {
        message: "Attribute not found",
      });
    }

    let updatedData = { ...req.body };
    // Update the category in the database
    const updatedAttribute = await Attribute.findByIdAndUpdate(
      id,
      updatedData,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Attribute updated successfully!",
      data: updatedAttribute,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

attributeController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    console.log(id);
    // Find the category by ID
    const attributeItem = await Attribute.findById(id);
    if (!attributeItem) {
      return sendResponse(res, 404, "Failed", {
        message: "Attribute not found",
      });
    }
    // Delete the category from the database
    await Attribute.findByIdAndDelete(id);

    sendResponse(res, 200, "Success", {
      message: "Attribute deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

module.exports = attributeController;
