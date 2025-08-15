const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const notifyController = express.Router();
const Notify = require("../model/notify.Schema");
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const { sendNotification } = require("../utils/sendNotification");


notifyController.post("/create", upload.single("icon"), async (req, res) => {
    try {
      let obj = req.body;
  
      if (req.file) {
        const icon = await cloudinary.uploader.upload(req.file.path);
        obj.icon = icon.url;
      }
  
      const notifyCreated = await Notify.create(obj);
  
      sendResponse(res, 200, "Success", {
        message: "notify created successfully!",
        data: notifyCreated,
        statusCode: 200
      });
  
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500
      });
    }
  });

notifyController.post("/list", async (req, res) => {
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
    const notifyList = await Notify.find(query)
      
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
   
    sendResponse(res, 200, "Success", {
      message: "Notify list retrieved successfully!",
      data: notifyList,
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

notifyController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const notify = await Notify.findById(id);
    if (!notify) {
      return sendResponse(res, 404, "Failed", {
        message: "Notify not found",
        statusCode: 403,
      });
    }
    const updatedNotify = await Notify.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, // Return the updated document
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Mark as read!",
      data: updatedNotify,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});


module.exports = notifyController;
