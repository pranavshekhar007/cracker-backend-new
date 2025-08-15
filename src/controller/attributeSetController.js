const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const AttributeSet = require("../model/attributeSet.Schema");
const Attribute = require("../model/attribute.Schema");
const attributeSetController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

attributeSetController.post("/create", async (req, res) => {
  try {
    const AttributeSetCreated = await AttributeSet.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Attribute Set created successfully!",
      data: AttributeSetCreated,
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

attributeSetController.post("/list", async (req, res) => {
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
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };
    const attributeSetList = await AttributeSet.find(query)
      .populate({
        path: "subCategoryId",
        
      })
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await AttributeSet.countDocuments({});
    const activeCount = await AttributeSet.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Attribute set list retrieved successfully!",
      data: attributeSetList,
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
      statusCode: 500,
    });
  }
});

attributeSetController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const attributeSet = await AttributeSet.findById(id);
    if (!attributeSet) {
      return sendResponse(res, 404, "Failed", {
        message: "Attribute set not found",
        statusCode: 403,
      });
    }
    const updatedAttributeSet = await AttributeSet.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Attribute set updated successfully!",
      data: updatedAttributeSet,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

attributeSetController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const attributeSet = await AttributeSet.findById(id);
    if (!attributeSet) {
      return sendResponse(res, 404, "Failed", {
        message: "Attribute set not found",
        statusCode: 404,
      });
    }
    await AttributeSet.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Attribute set deleted successfully!",
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

attributeSetController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const attributeSetDetails = await AttributeSet.findOne({ _id: id });
    const attributeList = await Attribute.find({ attributeSetId: id });
    sendResponse(res, 200, "Success", {
      message: "Attributeset retrived successfully!",
      data: { attributeSetDetails, attributeList },
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

module.exports = attributeSetController;
