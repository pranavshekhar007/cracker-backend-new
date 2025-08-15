const express = require("express");
const subscriptionChitController = express.Router();
const { sendResponse } = require("../utils/common");
const SubscriptionChit = require("../model/subscriptionChit.Schema");
const upload = require("../utils/multer");
const cloudinary = require("../utils/cloudinary");
require("dotenv").config();
const SchemeConfig = require("../model/schemeConfig.Schema");
const bcrypt = require("bcrypt");
const { generateRandomPassword } = require("../utils/password");
const sendEmail = require("../utils/sendEmail");
const User = require("../model/user.Schema");

// subscriptionChitController.post("/create", async (req, res) => {
//   try {
//     const { userId, chitSubscriptionSignUpId, totalAmount } = req.body;

//     const chit = await SubscriptionChit.findById(chitSubscriptionSignUpId);
//     if (!chit) {
//       return sendResponse(res, 404, "Failed", {
//         message: "Subscription chit not found",
//         statusCode: 404,
//       });
//     }

//     const config = await SchemeConfig.findOne();
//     if (!config) {
//       return sendResponse(res, 404, "Failed", {
//         message: "Scheme config not found. Contact admin.",
//         statusCode: 404,
//       });
//     }

//     const enrolmentDate = new Date();
//     const start = config.schemeStartDate;
//     const end = config.schemeEndDate;

//     const monthsDiff =
//       (end.getFullYear() - enrolmentDate.getFullYear()) * 12 +
//       (end.getMonth() - enrolmentDate.getMonth()) +
//       1;

//     const monthlyAmount = Math.ceil(totalAmount / monthsDiff);

//     chit.userId = userId;
//     chit.totalAmount = totalAmount;
//     chit.monthlyAmount = monthlyAmount;
//     chit.totalMonths = monthsDiff;
//     chit.schemeStartDate = start;
//     chit.schemeEndDate = end;
//     chit.enrolmentDate = enrolmentDate;

//     await chit.save();

//     sendResponse(res, 200, "Success", {
//       message: "Subscription chit purchased successfully!",
//       data: chit,
//       statusCode: 200,
//     });
//   } catch (error) {
//     console.error(error);
//     sendResponse(res, 500, "Failed", {
//       message: error.message || "Internal server error",
//       statusCode: 500,
//     });
//   }
// });

/**
 * List chit subscriptions
 */
// subscriptionChitController.post("/list", async (req, res) => {
//   try {
//     const {
//       searchKey = "",
//       userId,
//       pageNo = 1,
//       pageCount = 10,
//       sortByField,
//       sortByOrder,
//       status,
//     } = req.body;

//     const query = {};
//     if (userId) query.userId = userId;
//     if (status !== undefined && status !== "") query.status = status;

//     const sortField = sortByField || "createdAt";
//     const sortOrder = sortByOrder === "asc" ? 1 : -1;
//     const sortOption = { [sortField]: sortOrder };

//     const chitList = await SubscriptionChit.find(query)
//       .sort(sortOption)
//       .limit(parseInt(pageCount))
//       .skip((parseInt(pageNo) - 1) * parseInt(pageCount))
//       .populate("userId");

//     const totalCount = await SubscriptionChit.countDocuments(query);
//     const activeCount = await SubscriptionChit.countDocuments({
//       ...query,
//       status: true,
//     });

//     sendResponse(res, 200, "Success", {
//       message: "Subscription chit list retrieved successfully!",
//       data: chitList,
//       documentCount: {
//         totalCount,
//         activeCount,
//         inactiveCount: totalCount - activeCount,
//       },
//       statusCode: 200,
//     });
//   } catch (error) {
//     console.error(error);
//     sendResponse(res, 500, "Failed", {
//       message: error.message || "Internal server error",
//       statusCode: 500,
//     });
//   }
// });

/**
 * Update chit subscription or payment approval
 */

subscriptionChitController.put(
  "/update",
  upload.single("paymentSs"),
  async (req, res) => {
    try {
      const { _id, monthNumber, monthYear, paymentDate, action } = req.body;

      const chit = await SubscriptionChit.findById(_id);
      if (!chit) {
        return sendResponse(res, 404, "Failed", {
          message: "Subscription chit not found",
          statusCode: 404,
        });
      }

      let uploadedPaymentSsURL = "";
      if (req.file) {
        const uploaded = await cloudinary.uploader.upload(req.file.path);
        uploadedPaymentSsURL = uploaded.url;
      }

      // Convert numeric monthNumber to month name if needed
      let monthName = monthNumber;
      if (!isNaN(monthNumber)) {
        const date = new Date();
        date.setMonth(parseInt(monthNumber) - 1);
        monthName = date.toLocaleString("default", { month: "long" });
      }

      // Check if the month + year entry already exists
      const monthIndex = chit.paidMonths.findIndex(
        (m) => m.monthNumber === monthName && m.monthYear === monthYear
      );

      if (monthIndex >= 0) {
        // Update existing
        chit.paidMonths[monthIndex].paymentDate =
          paymentDate || chit.paidMonths[monthIndex].paymentDate;
        chit.paidMonths[monthIndex].status =
          action || chit.paidMonths[monthIndex].status;

        if (uploadedPaymentSsURL) {
          chit.paidMonths[monthIndex].paymentSs = uploadedPaymentSsURL;
        }
      } else {
        // Push new month entry
        chit.paidMonths.push({
          monthNumber: monthName,
          monthYear: monthYear || new Date().getFullYear().toString(),
          paymentDate,
          status: action || "pending",
          paymentSs: uploadedPaymentSsURL || "",
        });
      }

      await chit.save();

      sendResponse(res, 200, "Success", {
        message: "Subscription chit updated successfully!",
        data: chit,
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

/**
 * Get chit subscription details by ID
 */
subscriptionChitController.get("/details/:id", async (req, res) => {
  try {
    const id = req.params.id;

    const chit = await SubscriptionChit.findById(id).populate("userId");
    if (!chit) {
      return sendResponse(res, 404, "Failed", {
        message: "Subscription chit not found",
        statusCode: 404,
      });
    }

    sendResponse(res, 200, "Success", {
      message: "Subscription chit details retrieved successfully!",
      data: chit,
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

/**
 * Delete chit subscription by ID
 */
subscriptionChitController.delete("/delete/:id", async (req, res) => {
  try {
    const id = req.params.id;

    const chit = await SubscriptionChit.findById(id);
    if (!chit) {
      return sendResponse(res, 404, "Failed", {
        message: "Subscription chit not found",
        statusCode: 404,
      });
    }

    await SubscriptionChit.findByIdAndDelete(id);

    sendResponse(res, 200, "Success", {
      message: "Subscription chit deleted successfully!",
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

subscriptionChitController.put("/update/payment-status", async (req, res) => {
  try {
    const { chitId, monthNumber, status } = req.body;

    const chit = await SubscriptionChit.findById(chitId);
    if (!chit) {
      return sendResponse(res, 404, "Failed", {
        message: "Subscription chit not found",
        statusCode: 404,
      });
    }

    const monthIndex = chit.paidMonths.findIndex(
      (m) => m.monthNumber == monthNumber
    );

    if (monthIndex >= 0) {
      chit.paidMonths[monthIndex].status =
        status || chit.paidMonths[monthIndex].status;
    } else {
      return sendResponse(res, 404, "Failed", {
        message: "Month not found in paidMonths",
        statusCode: 404,
      });
    }

    await chit.save();

    sendResponse(res, 200, "Success", {
      message: `Month ${monthNumber} status updated to ${status}`,
      data: chit,
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

// Chit-Subscription Sign Up
subscriptionChitController.post("/signup", async (req, res) => {
  try {
    const { userId, name, phone, email, location, totalAmount } = req.body;

    if (!location) {
      return sendResponse(res, 400, "Failed", {
        message: "Location is required",
        statusCode: 400,
      });
    }

    let finalName = name;
    let finalEmail = email;
    let finalPhone = phone;

    if (userId) {
      const userData = await User.findById(userId);

      if (userData) {
        finalName = finalName || `${userData.firstName} ${userData.lastName}`;
        finalEmail = finalEmail || userData.email;
        finalPhone = finalPhone || userData.phone;
      }
    }

    if (!finalName || !finalEmail || !finalPhone) {
      return sendResponse(res, 400, "Failed", {
        message: "Name, Email, and Phone are required",
        statusCode: 400,
      });
    }

    // Check duplicate email
    const chitExists = await SubscriptionChit.findOne({ email: finalEmail });
    if (chitExists) {
      return sendResponse(res, 400, "Failed", {
        message: "User already registered with this email",
        statusCode: 400,
      });
    }

    // Fetch scheme config
    const config = await SchemeConfig.findOne();
    if (!config) {
      return sendResponse(res, 404, "Failed", {
        message: "Scheme config not found. Contact admin.",
        statusCode: 404,
      });
    }

    const enrolmentDate = new Date();
    const start = config.schemeStartDate;
    const end = config.schemeEndDate;

    let calculationStartMonth = enrolmentDate.getMonth(); // default: current month

    if (enrolmentDate.getDate() > 10) {
      calculationStartMonth += 1; // if enrolled after 10th, start from next month
    }

    // Adjust year if month exceeds December
    let calculationStartYear = enrolmentDate.getFullYear();
    if (calculationStartMonth > 11) {
      calculationStartMonth = 0;
      calculationStartYear += 1;
    }

    // Calculate total months from calculationStartMonth to scheme end
    const monthsDiff =
      (end.getFullYear() - calculationStartYear) * 12 +
      (end.getMonth() - calculationStartMonth) +
      1;

    if (monthsDiff <= 0) {
      return sendResponse(res, 400, "Failed", {
        message:
          "Scheme end date is before the calculated start month. Please check scheme configuration.",
        statusCode: 400,
      });
    }

    const monthlyAmount = Math.ceil(totalAmount / monthsDiff);

    // Create chit subscription with calculated fields
    const chitCreated = await SubscriptionChit.create({
      userId,
      name: finalName,
      email: finalEmail,
      phone: finalPhone,
      location,
      totalAmount,
      monthlyAmount,
      totalMonths: monthsDiff,
      schemeStartDate: start,
      schemeEndDate: end,
      enrolmentDate: enrolmentDate,
      status: "pending",
    });

    sendResponse(res, 200, "Success", {
      message: "Registration successful! Await admin approval.",
      data: chitCreated,
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

subscriptionChitController.put("/approve/:id", async (req, res) => {
  try {
    const { id } = req.params;

    const chit = await SubscriptionChit.findById(id);
    if (!chit) {
      return sendResponse(res, 404, "Failed", {
        message: "Subscription chit not found",
        statusCode: 404,
      });
    }

    if (chit.status === "approved") {
      return sendResponse(res, 400, "Failed", {
        message: "This user is already approved",
        statusCode: 400,
      });
    }

    // Generate random password
    const randomPassword = generateRandomPassword(10);

    // Hash password
    const salt = await bcrypt.genSalt(10);
    const hashedPassword = await bcrypt.hash(randomPassword, salt);

    // Update chit subscription
    chit.password = hashedPassword;
    chit.status = "approved";
    await chit.save();

    // Send email to user
    const subject = "Your Subscription Chit Account Approved";
    const html = `
      <p>Dear ${chit.name},</p>
  <p>Your subscription chit account has been approved.</p>
  <p><strong>Email:</strong> ${chit.email}</p>
  <p><strong>Password:</strong> ${randomPassword}</p>
  <hr/>
  <p><strong>Enrollment Date:</strong> ${chit.enrolmentDate.toDateString()}</p>
  <p><strong>Monthly Payment Amount:</strong> ₹${chit.monthlyAmount}</p>
  <p><strong>Scheme Duration:</strong> ${chit.totalMonths} months</p>
  <hr/>
  <p>You can now login using these credentials and start your monthly payments.</p>
  <p>Thank you!</p>
    `;

    await sendEmail(chit.email, subject, html);

    sendResponse(res, 200, "Success", {
      message: "User approved and credentials sent via email.",
      data: chit,
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

/**
 * Chit subscription login
 */
subscriptionChitController.post("/login", async (req, res) => {
  try {
    const { email, password } = req.body;

    const chit = await SubscriptionChit.findOne({ email });
    if (!chit) {
      return sendResponse(res, 404, "Failed", {
        message: "User not found",
        statusCode: 404,
      });
    }

    if (chit.status !== "approved") {
      return sendResponse(res, 401, "Failed", {
        message: "Your account is not approved yet.",
        statusCode: 401,
      });
    }

    const isMatch = await bcrypt.compare(password, chit.password);
    if (!isMatch) {
      return sendResponse(res, 401, "Failed", {
        message: "Invalid credentials",
        statusCode: 401,
      });
    }

    sendResponse(res, 200, "Success", {
      message: "Login successful",
      data: chit,
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

/**
 * List all registered chit subscription users with optional status filter
 */
subscriptionChitController.post(
  "/subscription-users/list",
  async (req, res) => {
    try {
      const {
        userId,
        searchKey = "",
        pageNo = 1,
        pageCount = 10,
        status,
      } = req.body;

      const query = {};
      if (userId) query.userId = userId;
      if (status) query.status = status;

      if (searchKey) {
        query.$or = [
          { name: { $regex: searchKey, $options: "i" } },
          { location: { $regex: searchKey, $options: "i" } },
        ];
      }

      const totalCount = await SubscriptionChit.countDocuments(query);

      const chitUsers = await SubscriptionChit.find(query)
        .sort({ createdAt: -1 })
        .skip((parseInt(pageNo) - 1) * parseInt(pageCount))
        .limit(parseInt(pageCount))
        .populate("userId");

      sendResponse(res, 200, "Success", {
        message: "Subscription chit users list retrieved successfully",
        data: chitUsers,
        totalCount,
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

/**
 * List subscriptions purchased by a chit user
 */
subscriptionChitController.post("/my-subscriptions", async (req, res) => {
  try {
    const { userId } = req.body;

    if (!userId) {
      return sendResponse(res, 400, "Failed", {
        message: "User ID is required",
        statusCode: 400,
      });
    }

    const subscriptions = await SubscriptionChit.find({ _id: userId }).populate(
      "userId"
    );

    sendResponse(res, 200, "Success", {
      message: "Subscriptions retrieved successfully",
      data: subscriptions,
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

module.exports = subscriptionChitController;
