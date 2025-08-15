const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const ProductType = require("../model/productType.Schema");
const productTypeController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

productTypeController.post("/create", async (req, res) => {
  try {
    const productTypeCreated = await ProductType.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Product Type created successfully!",
      data: productTypeCreated,
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

productTypeController.post("/list", async (req, res) => {
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
    const productTypeList = await ProductType.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await ProductType.countDocuments({});
    const activeCount = await ProductType.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Product type retrieved successfully!",
      data: productTypeList,
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

productTypeController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const productType = await ProductType.findById(id);
    if (!productType) {
      return sendResponse(res, 404, "Failed", {
        message: "Product type not found",
        statusCode: 403,
      });
    }
    const updatedProductType = await ProductType.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Product type updated successfully!",
      data: updatedProductType,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

productTypeController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const productType = await ProductType.findById(id);
    if (!productType) {
      return sendResponse(res, 404, "Failed", {
        message: "Product type not found",
        statusCode: 404,
      });
    }
    await ProductType.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Product type deleted successfully!",
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

productTypeController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const productTypeDetails = await ProductType.findOne({ _id: id });
    sendResponse(res, 200, "Success", {
      message: "Product type retrived successfully!",
      data: { productTypeDetails },
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

module.exports = productTypeController;
