const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Tag = require("../model/tag.Schema");
const tagController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

tagController.post("/create", async (req, res) => {
  try {
    const tagCreated = await Tag.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Tag created successfully!",
      data: tagCreated,
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

tagController.post("/list", async (req, res) => {
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
    const tagList = await Tag.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await Tag.countDocuments(query);
    const activeCount = await Tag.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Tag list retrieved successfully!",
      data: tagList,
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

tagController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const tag = await Tag.findById(id);
    if (!tag) {
      return sendResponse(res, 404, "Failed", {
        message: "Tag set not found",
        statusCode: 403,
      });
    }
    const updatedTag = await Tag.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Tag updated successfully!",
      data: updatedTag,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

tagController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const tag = await Tag.findById(id);
    if (!tag) {
      return sendResponse(res, 404, "Failed", {
        message: "Tag not found",
        statusCode: 404,
      });
    }
    await Tag.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Tag deleted successfully!",
      statusCode:200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,   
    });
  }
});

tagController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const tagDetails = await Tag.findOne({ _id: id });
    sendResponse(res, 200, "Success", {
      message: "Tag retrived successfully!",
      data: { tagDetails },
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

module.exports = tagController;
