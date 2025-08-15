const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Role = require("../model/role.Schema");
const roleController = express.Router();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

roleController.post("/create", async (req, res) => {
  try {
    const roleCreated = await Role.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Role created successfully!",
      data: roleCreated,
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

roleController.post("/list", async (req, res) => {
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
    const roleList = await Role.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await Role.countDocuments({});
    const activeCount = await Role.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Role list retrieved successfully!",
      data: roleList,
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

roleController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const role = await Role.findById(id);
    if (!role) {
      return sendResponse(res, 404, "Failed", {
        message: "Role not found",
        statusCode: 403,
      });
    }
    const updatedRole = await Role.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Role updated successfully!",
      data: updatedRole,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

roleController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const role = await Role.findById(id);
    if (!role) {
      return sendResponse(res, 404, "Failed", {
        message: "Role not found",
        statusCode: 404,
      });
    }
    await Role.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Role deleted successfully!",
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

module.exports = roleController;