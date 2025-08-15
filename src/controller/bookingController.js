const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Booking = require("../model/booking.Schema");
const bookingController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const auth = require("../utils/auth");
const fs = require("fs");
const path = require("path");
const sendEmail = require("../utils/sendEmail");
const Address = require("../model/address.Schema");
const Area = require("../model/area.Schema");
const Coupon = require("../model/coupon.Schema");
const User = require("../model/user.Schema");

bookingController.post("/create", async (req, res) => {
  try {
    const bookingData = {
      ...req.body,
    };

    let product = req.body.product.filter((v, i) => v.productType == "Product");
    let comboProduct = req.body.product
      .filter((v) => v?.productType === "ComboProduct")
      .map((v) => ({
        comboProductId: v.productId,
        quantity: v.quantity,
        totalPrice: v.totalPrice,
      }));

    // let couponId = req.body.couponId;

    // if (couponId) {
    //   const coupon = await Coupon.findOne({ _id: couponId });

    //   if (!coupon) {
    //     return sendResponse(res, 400, "Failed", {
    //       message: "Invalid coupon code",
    //       statusCode: 400,
    //     });
    //   }

    //   const now = new Date();

    //   let totalAmount = req.body.totalAmount;

    //   const isValid =
    //     coupon.status === "active" &&
    //     now >= new Date(coupon.validFrom) &&
    //     now <= new Date(coupon.validTo) &&
    //     totalAmount >= coupon.minimumOrderAmount;

    //   if (!isValid) {
    //     return sendResponse(res, 400, "Failed", {
    //       message:
    //         "Coupon is not valid at this time or order amount is too low.",
    //       statusCode: 400,
    //     });
    //   }
    //   if (coupon.usedCount == coupon.usageLimit) {
    //     return sendResponse(res, 400, "Failed", {
    //       message: "Coupon is not valid, reach the maximum use.",
    //       statusCode: 400,
    //     });
    //   }
    //   await Coupon.findByIdAndUpdate(
    //     couponId,
    //     { $set: { usedCount: coupon.usedCount + 1 } },
    //     { new: true }
    //   );
    // }

    const bookingCreated = await Booking.create({
      ...bookingData,
      product,
      comboProduct,
    });

    const populatedBooking = await Booking.findById(bookingCreated._id)
      .populate("product.productId")
      .populate({
        path: "comboProduct.comboProductId",
        populate: {
          path: "productId.product",
          model: "Product",
        },
      });

      await User.findByIdAndUpdate(bookingData.userId, {
        $set: { cartItems: [] },
      });      

    sendResponse(res, 200, "Success", {
      message: "Booking created successfully!",
      data: populatedBooking,
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

bookingController.post("/list", async (req, res) => {
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
    if (searchKey) query.status = { $regex: searchKey, $options: "i" };

    // Construct sorting object
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };

    // Fetch the booking list
    const bookingList = await Booking.find(query)
      .populate("userId", "firstName lastName")
      .populate("cityId", "name minimumPrice")
      .populate({
        path: "product.productId",
        select: "name description productHeroImage",
      })
      .populate({
        path: "comboProduct.comboProductId",
        select: "name productHeroImage pricing productId",
        populate: {
          path: "productId.product",
          model: "Product",
          select: "name price productHeroImage",
        },
      })      
      .sort(sortOption)
      .skip(parseInt(pageNo - 1) * parseInt(pageCount))
      .limit(parseInt(pageCount));

    const totalCount = await Booking.countDocuments(query);

    const statusCounts = await Booking.aggregate([
      { $match: query },
      { $group: { _id: "$status", count: { $sum: 1 } } },
    ]);

    // Map status counts
    const statusCountMap = {
      orderPlaced: 0,
      orderPacked: 0,
      orderDispatch: 0,
      completed: 0,
      cancelled: 0,
      shipping: 0,
      pending: 0,
    };

    statusCounts.forEach(({ _id, count }) => {
      statusCountMap[_id] = count;
    });

    //  Date calculations
    const todayStart = new Date();
    todayStart.setHours(0, 0, 0, 0);

    const weekStart = new Date();
    weekStart.setDate(weekStart.getDate() - weekStart.getDay());
    weekStart.setHours(0, 0, 0, 0);

    const monthStart = new Date();
    monthStart.setDate(1);
    monthStart.setHours(0, 0, 0, 0);

    // Counts based on createdAt date
    const todaysOrder = await Booking.countDocuments({
      ...query,
      createdAt: { $gte: todayStart },
    });

    const thisWeekOrder = await Booking.countDocuments({
      ...query,
      createdAt: { $gte: weekStart },
    });

    const thisMonthOrder = await Booking.countDocuments({
      ...query,
      createdAt: { $gte: monthStart },
    });

    sendResponse(res, 200, "Success", {
      message: "Booking list retrieved successfully!",
      data: bookingList,
      documentCount: {
        totalCount,
        ...statusCountMap,
        todaysOrder,
        thisWeekOrder,
        thisMonthOrder,
      },
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

bookingController.get("/details/:id", async (req, res) => {
  try {
    const id = req.params.id;
    const booking = await Booking.findOne({ _id: id })
      .populate("userId", "firstName lastName email phone")
      .populate("product.productId")
      .populate({
        path: "comboProduct.comboProductId",
        populate: {
          path: "productId.product",
          model: "Product",
          select: "name price productHeroImage",
        },
      })
      .lean();

    if (!booking) {
      return sendResponse(res, 404, "Failed", {
        message: "No bookings found",
        statusCode: 404,
      });
    }

    if (booking.address?.area) {
      const areaData = await Area.findOne({
        name: booking.address.area,
      }).lean();
      if (areaData) {
        booking.address.area = {
          _id: areaData._id,
          name: areaData.name,
        };
      } else {
        booking.address.area = { name: booking.address.area }; // fallback
      }
    }

    return sendResponse(res, 200, "Success", {
      message: "Booking details fetched successfully",
      data: booking,
      statusCode: 200,
    });
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

bookingController.get("/user/:userId", async (req, res) => {
  try {
    console.log("Hello :", req.params.userId);
    const userId = req.params.userId;
    const booking = await Booking.find({ userId: userId })
      .populate("product.productId")
      .populate("userId")
      .populate({
        path: "comboProduct.comboProductId",
        populate: {
          path: "productId.product",
          model: "Product",
          select: "name price productHeroImage",
        },
      });

    if (booking.length > 0) {
      return sendResponse(res, 200, "Success", {
        message: "Booking details fetched successfully",
        data: booking,
        statusCode: 200,
      });
    } else {
      return sendResponse(res, 404, "Failed", {
        message: "No bookings found for this user",
        statusCode: 404,
      });
    }
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

bookingController.put("/update", async (req, res) => {
  try {
    const { id, status, ...updateFields } = req.body;

    if (!id) {
      return sendResponse(res, 400, "Failed", {
        message: "Booking ID is required",
        statusCode: 400,
      });
    }

    const booking = await Booking.findById(id).populate("userId");
    if (!booking) {
      return sendResponse(res, 404, "Failed", {
        message: "Booking not found",
        statusCode: 404,
      });
    }

    // Update the statusHistory if the status has changed
    if (status && status !== booking.status) {
      if (!booking.statusHistory) booking.statusHistory = [];

      booking.statusHistory.push({
        status: status,
        updatedAt: new Date(),
      });
    }

    // Merge updated fields including status
    const updatedFields = {
      ...updateFields,
      ...(status && { status }), // only add status if it exists
      statusHistory: booking.statusHistory,
    };

    const updatedBooking = await Booking.findByIdAndUpdate(id, updatedFields, {
      new: true,
    })
      .populate("userId", "firstName lastName email")
      .populate("product.productId", "name description productHeroImage");

    // ✅ Fetch address email and send email if status is orderPlaced
    if (status === "orderPlaced") {
      let userEmail = "";
      let userName = updatedBooking.userId.firstName;

      // Fetch the Address document linked to this booking
      const addressData = await Address.findOne({
        userId: updatedBooking.userId._id,
      });

      if (addressData) {
        userEmail = addressData.email;
      } else {
        userEmail = updatedBooking.userId.email; // fallback to user email if address not found
      }

      const html = `
        <p>Dear ${userName},</p>
        <p>Your order with Booking ID <b>${updatedBooking._id}</b> has been placed successfully.</p>
        <p>Thank you for shopping with us!</p>
      `;

      await sendEmail(userEmail, "Your Order is Placed!", html);
    }

    return sendResponse(res, 200, "Success", {
      message: "Booking updated successfully",
      data: updatedBooking,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

bookingController.put(
  "/upload/payment-ss",
  upload.single("paymentSs"),
  async (req, res) => {
    try {
      const bookingId = req.body.id;

      const booking = await Booking.findById(bookingId);
      if (!booking) {
        return sendResponse(res, 404, "Failed", {
          message: "Booking not found",
          statusCode: 404,
        });
      }

      let updatedData = { ...req.body };

      if (req.file) {
        const paymentScreenshot = await cloudinary.uploader.upload(
          req.file.path
        );
        updatedData.paymentSs = paymentScreenshot.url;
      }

      const updatedBooking = await Booking.findByIdAndUpdate(
        bookingId,
        updatedData,
        { new: true }
      );

      sendResponse(res, 200, "Success", {
        message: "Payment screenshot uploaded successfully!",
        data: updatedBooking,
        statusCode: 200,
      });
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500,
      });
    }
  }
);

module.exports = bookingController;
