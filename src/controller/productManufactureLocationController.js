const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const ProductManufactureLocation = require("../model/productManufactureLocation.Schema");
const productManufactureLocationController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

productManufactureLocationController.post("/create", async (req, res) => {
  try {
    const productLocationCreated = await ProductManufactureLocation.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Product Manufacture Location created successfully!",
      data: productLocationCreated,
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

productManufactureLocationController.post("/list", async (req, res) => {
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
    const productLocation = await ProductManufactureLocation.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await ProductManufactureLocation.countDocuments({});
    const activeCount = await ProductManufactureLocation.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Product Manufacture Location retrieved successfully!",
      data: productLocation,
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

productManufactureLocationController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const productLocation = await ProductManufactureLocation.findById(id);
    if (!productLocation) {
      return sendResponse(res, 404, "Failed", {
        message: "Product location not found",
        statusCode: 403,
      });
    }
    const updatedProductLocation = await ProductManufactureLocation.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Product location updated successfully!",
      data: updatedProductLocation,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

productManufactureLocationController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const productLocation = await ProductManufactureLocation.findById(id);
    if (!productLocation) {
      return sendResponse(res, 404, "Failed", {
        message: "Product location not found",
        statusCode: 404,
      });
    }
    await ProductManufactureLocation.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Product location deleted successfully!",
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

productManufactureLocationController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const productLocationDetails = await ProductManufactureLocation.findOne({ _id: id });
    sendResponse(res, 200, "Success", {
      message: "Product location retrived successfully!",
      data:  productLocationDetails ,
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

module.exports = productManufactureLocationController;
