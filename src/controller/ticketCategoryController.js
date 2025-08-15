const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const   TicketCategory = require("../model/ticketCategory.Schema");
const ticketCategoryController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

ticketCategoryController.post("/create", async (req, res) => {
  try {
    const ticketCategoryCreated = await TicketCategory.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Ticket Category created successfully!",
      data: ticketCategoryCreated,
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

ticketCategoryController.post("/list", async (req, res) => {
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
    const ticketCategoryList = await TicketCategory.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await TicketCategory.countDocuments({});
    const activeCount = await TicketCategory.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "Ticket Category list retrieved successfully!",
      data: ticketCategoryList,
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

ticketCategoryController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const ticketCategory = await TicketCategory.findById(id);
    if (!ticketCategory) {
      return sendResponse(res, 404, "Failed", {
        message: "Ticket Category not found",
        statusCode: 403,
      });
    }
    const updatedTicketCategory = await TicketCategory.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Ticket Category updated successfully!",
      data: updatedTicketCategory,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

ticketCategoryController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const ticketCategory = await TicketCategory.findById(id);
    if (!ticketCategory) {
      return sendResponse(res, 404, "Failed", {
        message: "Ticket Category not found",
        statusCode: 404,
      });
    }
    await TicketCategory.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Ticket Category deleted successfully!",
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

ticketCategoryController.get("/details/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const ticketCategoryDetails = await TicketCategory.findOne({ _id: id });
    sendResponse(res, 200, "Success", {
      message: "Ticket Category retrived successfully!",
      data: { ticketCategoryDetails },
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

module.exports = ticketCategoryController;
