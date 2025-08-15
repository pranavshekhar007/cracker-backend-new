const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const   Tax = require("../model/tax.Schema");
const taxController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

taxController.post("/create", async (req, res) => {
  try {
    const taxCreated = await Tax.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Tax created successfully!",
      data: taxCreated,
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

taxController.post("/list", async (req, res) => {
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
    const taxList = await Tax.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await Tax.countDocuments({});
    const activeCount = await Tax.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "tax list retrieved successfully!",
      data: taxList,
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

taxController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const tax = await Tax.findById(id);
    if (!tax) {
      return sendResponse(res, 404, "Failed", {
        message: "Tax not found",
        statusCode: 403,
      });
    }
    const updatedtax = await Tax.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Tax updated successfully!",
      data: updatedtax,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

taxController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const tax = await Tax.findById(id);
    if (!tax) {
      return sendResponse(res, 404, "Failed", {
        message: "Tax not found",
        statusCode: 404,
      });
    }
    await Tax.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Tax deleted successfully!",
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

taxController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const taxDetails = await Tax.findOne({ _id: id });
    sendResponse(res, 200, "Success", {
      message: "Tax retrived successfully!",
      data: { taxDetails },
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

module.exports = taxController;
