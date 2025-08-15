const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Permission = require("../model/permission.Schema");
const permissionController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

permissionController.post("/create", async (req, res) => {
  try {
    const permissionCreated = await Permission.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Permission created successfully!",
      data: permissionCreated,
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

permissionController.post("/list", async (req, res) => {
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
    const permissionList = await Permission.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await Permission.countDocuments({});
    const activeCount = await Permission.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Permission list retrieved successfully!",
      data: permissionList,
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

permissionController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const permission = await Permission.findById(id);
    if (!permission) {
      return sendResponse(res, 404, "Failed", {
        message: "Permission not found",
        statusCode: 403,
      });
    }
    const updatedPermission = await Permission.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Permission updated successfully!",
      data: updatedPermission,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

permissionController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const permission = await Permission.findById(id);
    if (!permission) {
      return sendResponse(res, 404, "Failed", {
        message: "Permission not found",
        statusCode: 404,
      });
    }
    await Permission.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Permission deleted successfully!",
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

module.exports = permissionController;