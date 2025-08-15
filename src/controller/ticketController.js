const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Ticket = require("../model/ticket.Schema");
const User = require("../model/user.Schema");
const Driver = require("../model/driver.Schema");
const Vender = require("../model/vender.Schema");
const ticketController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");

ticketController.post("/create", upload.single("image"), async (req, res) => {
  try {
    let obj = req.body;

    if (req.file) {
      const image = await cloudinary.uploader.upload(req.file.path);
      obj.image = image.url;
    }

    const ticketCreated = await Ticket.create(obj);

    sendResponse(res, 200, "Success", {
      message: "Ticket created successfully!",
      data: ticketCreated,
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

ticketController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      status,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
      userId,
      userType
    } = req.body;

    const query = {};
    if (status !== undefined && status !== "") query.status = status;
    if (userType !== undefined && userType !== "") query.userType = userType;
    if (searchKey) query.subject = { $regex: searchKey, $options: "i" };
    if (userId) query.userId = userId;

    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };

    const ticketList = await Ticket.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip((parseInt(pageNo) - 1) * parseInt(pageCount))
      .populate("ticketCategoryId");

    // fetch user/vendor/driver/admin details based on userType
    const enhancedTickets = await Promise.all(
      ticketList.map(async (ticket) => {
        let userDetails = null;

        switch (ticket.userType) {
          case "User":
            userDetails = await User.findById(ticket.userId).lean();
            break;
          case "Vender":
            userDetails = await Vender.findById(ticket.userId).lean();
            break;
          case "Driver":
            userDetails = await Driver.findById(ticket.userId).lean();
            break;
          default:
            userDetails = null;
        }

        return {
          ...ticket.toObject(),
          userDetails,
        };
      })
    );

    const totalCount = await Ticket.countDocuments({userType:userType});
    const activeCount = await Ticket.countDocuments({ status: true, userType:userType });

    sendResponse(res, 200, "Success", {
      message: "Ticket list retrieved successfully!",
      data: enhancedTickets,
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

ticketController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const ticket = await Ticket.findById(id);
    if (!ticket) {
      return sendResponse(res, 404, "Failed", {
        message: "Ticket not found",
        statusCode: 403,
      });
    }
    const updatedTicket = await Ticket.findByIdAndUpdate(
      id,
      req.body,
      {
        new: true, 
      }
    );
    sendResponse(res, 200, "Success", {
      message: "Ticket updated successfully!",
      data: updatedTicket,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

module.exports = ticketController;
