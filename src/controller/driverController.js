const express = require("express");
const { sendResponse, generateOTP } = require("../utils/common");
require("dotenv").config();
const Driver = require("../model/driver.Schema");
const Booking = require("../model/booking.Schema");
const driverController = express.Router();
const axios = require("axios");
require("dotenv").config();
const jwt = require("jsonwebtoken");
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const { sendNotification } = require("../utils/sendNotification");
const auth = require("../utils/auth");


driverController.post(
  "/sign-up",
  upload.fields([
    { name: "dlFrontImage", maxCount: 1 },
    { name: "dlBackImage", maxCount: 1 },
    { name: "profilePic", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      // Check if the phone number is unique
      const user = await Driver.findOne({ phone: req.body.phone });
      if (user) {
        return sendResponse(res, 400, "Failed", {
          message: "Driver is already registered.",
          statusCode: 400,
        });
      }

      // Generate OTP
      const otp = generateOTP();

      // Upload images to Cloudinary
      let dlFrontImage, dlBackImage, profilePic;

      if (req.files["dlFrontImage"]) {
        let image = await cloudinary.uploader.upload(
          req.files["dlFrontImage"][0].path
        );
        dlFrontImage = image.url;
      }

      if (req.files["dlBackImage"]) {
        let image = await cloudinary.uploader.upload(
          req.files["dlBackImage"][0].path
        );
        dlBackImage = image.url;
      }
      if (req.files["profilePic"]) {
        let image = await cloudinary.uploader.upload(
          req.files["profilePic"][0].path
        );
        profilePic = image.url;
      }

      // Create a new user with provided details
      let newDriver = await Driver.create({
        ...req.body,
        phoneOtp: otp,
        dlFrontImage,
        dlBackImage,
        profilePic,
      });
      sendNotification({
        icon:newDriver.profilePic,
        title:"A new driver registered",
        subTitle:`${newDriver.firstName} has registered to the portal`,
        notifyUserId:"Admin",
        category:"Driver",
        subCategory:"Registration",
        notifyUser:"Admin",
      },req.io)

      // Generate JWT token
      const token = jwt.sign(
        { userId: newDriver._id, phone: newDriver.phone },
        process.env.JWT_KEY
      );

      // Store the token in the user object or return it in the response
      newDriver.token = token;
      const updatedDriver = await Driver.findByIdAndUpdate(
        newDriver._id,
        { token },
        { new: true }
      );

      // OTP message for autofill
      const appHash = "ems/3nG2V1H"; // Replace with your actual hash
      const otpMessage = `<#> ${otp} is your OTP for verification. Do not share it with anyone.\n${appHash}`;

      let otpResponse = await axios.post(
        `https://api.authkey.io/request?authkey=${
          process.env.AUTHKEY_API_KEY
        }&mobile=${req.body.phone}&country_code=91&sid=${
          process.env.AUTHKEY_SENDER_ID
        }&company=Acediva&otp=${otp}&message=${encodeURIComponent(otpMessage)}`
      );

      if (otpResponse?.status == "200") {
        return sendResponse(res, 200, "Success", {
          message: "OTP sent successfully",
          data: updatedDriver,
          statusCode: 200,
        });
      } else {
        return sendResponse(res, 422, "Failed", {
          message: "Unable to send OTP",
          statusCode: 200,
        });
      }
    } catch (error) {
      console.error("Error in /sign-up:", error.message);
      return sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error.",
      });
    }
  }
);

driverController.post("/otp-verification", async (req, res) => {
  try {
    const { phone, phoneOtp, isforgetPassword } = req.body;

    const user = await Driver.findOne({ phone, phoneOtp });
    if (user) {
      const updatedFields = {
        isPhoneVerified: true,
      };

      if (!isforgetPassword) {
        updatedFields.profileStatus = "completed";
      }

      const updatedDriver = await Driver.findByIdAndUpdate(
        user._id,
        updatedFields,
        { new: true }
      );

      sendNotification({
        icon: updatedDriver.profilePic,
        title: `${updatedDriver.firstName} has verified their phone number`,
        subTitle: `${updatedDriver.firstName} has verified their phone number`,
        notifyUserId: "Admin",
        category: "Driver",
        subCategory: "Verification",
        notifyUser: "Admin",
      });

      return sendResponse(res, 200, "Success", {
        message: "Otp verified successfully",
        data: updatedDriver,
        statusCode: 200,
      });
    } else {
      return sendResponse(res, 422, "Failed", {
        message: "Wrong OTP",
        statusCode: 422,
      });
    }
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

driverController.post("/login", async (req, res) => {
  try {
    const { phone, password } = req.body;
    const user = await Driver.findOne({ phone, password });
    if (user) {
      return sendResponse(res, 200, "Success", {
        message: "Driver logged in successfully",
        data: user,
        statusCode: 200,
      });
    } else {
      return sendResponse(res, 422, "Failed", {
        message: "Invalid Credentials",
        statusCode: 422,
      });
    }
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

driverController.post("/resend-otp", async (req, res) => {
  try {
    const { phone } = req.body;
    const user = await Driver.findOne({ phone });
    if (user) {
      const otp = generateOTP();
      const updatedDriver = await Driver.findByIdAndUpdate(
        user._id,
        { phoneOtp: otp },
        { new: true }
      );

      // OTP message for autofill
      const appHash = "ems/3nG2V1H"; // Replace with your actual hash
      const otpMessage = `<#> ${otp} is your OTP for verification. Do not share it with anyone.\n${appHash}`;

      let otpResponse = await axios.post(
        `https://api.authkey.io/request?authkey=${
          process.env.AUTHKEY_API_KEY
        }&mobile=${req.body.phone}&country_code=91&sid=${
          process.env.AUTHKEY_SENDER_ID
        }&company=Acediva&otp=${otp}&message=${encodeURIComponent(otpMessage)}`
      );

      if (otpResponse?.status == "200") {
        return sendResponse(res, 200, "Success", {
          message: "OTP sent successfully",
          data: updatedDriver,
          statusCode: 200,
        });
      } else {
        return sendResponse(res, 422, "Failed", {
          message: "Unable to send OTP",
          statusCode: 200,
        });
      }
    } else {
      return sendResponse(res, 422, "Failed", {
        message: "Phone number is not registered",
        statusCode: 422,
      });
    }
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

driverController.get("/details/:id", auth, async (req, res) => {
  try {
    const id = req.params.id;
    const driver = await Driver.findOne({ _id: id });
    if (driver) {
      return sendResponse(res, 200, "Success", {
        message: "Driver details fetched  successfully",
        data: driver,
        statusCode: 200,
      });
    } else {
      return sendResponse(res, 404, "Failed", {
        message: "Driver not found",
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

driverController.post("/list", auth, async (req, res) => {
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
    if (status) query.profileStatus = status;
    if (searchKey) {
      query.$or = [
        { firstName: { $regex: searchKey, $options: "i" } },
        { lastName: { $regex: searchKey, $options: "i" } },
        { email: { $regex: searchKey, $options: "i" } },
      ];
    }
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };
    const userList = await Driver.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await Driver.countDocuments({});
    const activeCount = await Driver.countDocuments({
      profileStatus: "approved",
    });
    sendResponse(res, 200, "Success", {
      message: "Driver list retrieved successfully!",
      data: userList,
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

driverController.post("/create", auth, async (req, res) => {
  try {
    const driver = await Driver.create(req.body);
    return sendResponse(res, 200, "Success", {
      message: "Driver created  successfully",
      data: driver,
      statusCode: 200,
    });
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

driverController.delete("/delete/:id", auth, async (req, res) => {
  try {
    const { id } = req.params;
    const driver = await Driver.findById(id);
    if (!driver) {
      return sendResponse(res, 404, "Failed", {
        message: "Driver not found",
      });
    }
    await Driver.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Driver deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

driverController.put("/update", auth, upload.fields([
    { name: "dlFrontImage", maxCount: 1 },
    { name: "dlBackImage", maxCount: 1 },
    { name: "profilePic", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const id = req.body.id;
      const userData = await Driver.findById(id);
      if (!userData) {
        return sendResponse(res, 404, "Failed", {
          message: "Driver not found",
        });
      }
      let updateData = {...req.body}
      if(req.file || req.files){
      
        if (req.files["dlFrontImage"]) {
          let image = await cloudinary.uploader.upload(
            req.files["dlFrontImage"][0].path
          );
          updateData = {...req.body, dlFrontImage: image.url};
        }
  
        if (req.files["dlBackImage"]) {
          let image = await cloudinary.uploader.upload(
            req.files["dlBackImage"][0].path
          );
          updateData = {...req.body, dlBackImage: image.url};
        }
        if (req.files["profilePic"]) {
          let image = await cloudinary.uploader.upload(
            req.files["profilePic"][0].path
          );
          updateData = {...req.body, profilePic: image.url};
        }
        
        const updatedUserData = await Driver.findByIdAndUpdate(id, updateData, {
          new: true, 
        });
        if(req.body.profileStatus=="reUploaded"){
          sendNotification({
            icon:updatedUserData.profilePic,
            title:`${updatedUserData.firstName} has re-uploaded the details`,
            subTitle:`${updatedUserData.firstName} has re-uploaded the details`,
            notifyUserId:"Admin",
            category:"Driver",
            subCategory:"Profile update",
            notifyUser:"Admin",
          }, req.io)
        }
        if(req.body.profileStatus=="rejected"){
          sendNotification({
            icon:updatedUserData.profilePic,
            title:`${updatedUserData.firstName} your details has been rejected`,
            subTitle:`${updatedUserData.firstName} please go through the details once more`,
            notifyUserId:updatedUserData._id,
            category:"Driver",
            subCategory:"Profile update",
            notifyUser:"Driver",
          }, req.io)
        }
        if(req.body.profileStatus=="approved"){
          sendNotification({
            icon:updatedUserData.profilePic,
            title:`${updatedUserData.firstName} your profile has been approved`,
            subTitle:`${updatedUserData.firstName} congratulations!! your profile has been approved`,
            notifyUserId:updatedUserData._id,
            category:"Driver",
            subCategory:"Profile update",
            notifyUser:"Driver",
          }, req.io)
        }
        sendResponse(res, 200, "Success", {
          message: "Driver updated successfully!",
          data: updatedUserData,
          statusCode: 200,
        });
      }
      else{
        const updatedUserData = await Driver.findByIdAndUpdate(id, updateData, {
          new: true, 
        });
  
        sendResponse(res, 200, "Success", {
          message: "Driver updated successfully!",
          data: updatedUserData,
          statusCode: 200,
        });

      }
      
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
      });
    }
  }
);

driverController.post("/assign-product", auth, async (req, res) => {
  try {
    const { orderId, productId, driverId } = req.body;

    const order = await Booking.findById(orderId);
    if (!order) {
      return sendResponse(res, 404, "Failed", {
        message: "Order not found",
        statusCode: 404,
      });
    }

    const driver = await Driver.findById(driverId);
    if (!driver) {
      return sendResponse(res, 404, "Failed", {
        message: "Driver not found",
        statusCode: 404,
      });
    }

    let productFound = false;

    order.product = order.product.map((item) => {
      if (item.productId.toString() === productId) {
        productFound = true;
        item.driverId = driverId;
      }
      return item;
    });

    if (!productFound) {
      return sendResponse(res, 404, "Failed", {
        message: "Product not found in order",
        statusCode: 404,
      });
    }

    await order.save();

    return sendResponse(res, 200, "Success", {
      message: "Product assigned to driver successfully",
      data: order,
      statusCode: 200,
    });
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
      statusCode: 500,
    });
  }
});

driverController.get("/assigned-products/:driverId", auth, async (req, res) => {
  try {
    const { driverId } = req.params;

    const orders = await Booking.find({ "product.driverId": driverId })
      .populate("product.productId")
      .populate("product.driverId")
      .populate("userId") 
      .populate("addressId"); 

    const assignedProducts = [];

    orders.forEach(order => {
      order.product.forEach(item => {
        if (item.driverId && item.driverId._id.toString() === driverId) {
          assignedProducts.push({
            orderId: order._id,
            product: item.productId,
            quantity: item.quantity,
            totalPrice: item.totalPrice,
            deliveryStatus: item.deliveryStatus,
            customer: order.userId,
            address: order.addressId,
            assignedAt: order.updatedAt
          });
        }
      });
    });

    return sendResponse(res, 200, "Success", {
      message: "Assigned products fetched successfully",
      data: assignedProducts,
      statusCode: 200,
    });
  } catch (error) {
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});


module.exports = driverController;
