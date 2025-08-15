const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const notificationController = express.Router();
const Notification = require("../model/notification.Schema");
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const { sendNotification } = require("../utils/sendNotification");


notificationController.post("/list", async (req, res) => {
  try {
    const {
      category,
      notifyUser,
      isRead,
      pageNo = 1,
      pageCount = 10,
    } = req.body;
    const query = {};
    if(category){
      query.category = category
    }
    if(notifyUser){
      query.notifyUser = notifyUser
    }
    if(isRead){
      query.isRead = isRead
    }
    const notificationList = await Notification.find(query)
      
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
   
    sendResponse(res, 200, "Success", {
      message: "Notification list retrieved successfully!",
      data: notificationList,
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

notificationController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const notification = await Notification.findById(id);
    if (!notification) {
      return sendResponse(res, 404, "Failed", {
        message: "Notification not found",
        statusCode: 403,
      });
    }
    const updatedNotifcation = await Notification.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Mark as read!",
      data: updatedNotifcation,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

notificationController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const notification = await Notification.findById(id);
    if (!notification) {
      return sendResponse(res, 404, "Failed", {
        message: "Notification not found",
        statusCode: 404,
      });
    }
    await Notification.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Notification deleted successfully!",
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

module.exports = notificationController;