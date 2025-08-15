const express = require("express");
const { sendResponse, generateOTP } = require("../utils/common");
require("dotenv").config();
const mongoose = require("mongoose");
const Product = require("../model/product.Schema");
const ComboProduct = require("../model/comboProduct.Schema");
const Category = require("../model/category.Schema");
const Banner = require("../model/banner.Schema");
const SubCategory = require("../model/subCategory.Schema");
const User = require("../model/user.Schema");
const Booking = require("../model/booking.Schema");
const userController = express.Router();
const axios = require("axios");
require("dotenv").config();
const jwt = require("jsonwebtoken");
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const auth = require("../utils/auth");
const moment = require("moment");
userController.post("/send-otp", async (req, res) => {
  try {
    const { phone, ...otherDetails } = req.body;
    // Check if the phone number is provided
    if (!phone) {
      return sendResponse(res, 400, "Failed", {
        message: "Phone number is required.",
        statusCode: 400,
      });
    }
    // Generate OTP
    const phoneOtp = generateOTP();

    // Check if the user exists
    let user = await User.findOne({ phone });

    if (!user) {
      // Create a new user with the provided details and OTP
      user = await User.create({
        phone,
        phoneOtp,
        ...otherDetails,
      });

      // Generate JWT token for the new user
      const token = jwt.sign(
        { userId: user._id, phone: user.phone },
        process.env.JWT_KEY
      );
      // Store the token in the user object or return it in the response
      user.token = token;
      user = await User.findByIdAndUpdate(user.id, { token }, { new: true });
    } else {
      // Update the existing user's OTP
      user = await User.findByIdAndUpdate(user.id, { phoneOtp }, { new: true });
    }
    const appHash = "ems/3nG2V1H"; // Apne app ka actual hash yahan dalein

    // Properly formatted OTP message for autofill
    const otpMessage = `<#> ${phoneOtp} is your OTP for verification. Do not share it with anyone.\n${appHash}`;

    let optResponse = await axios.post(
      `https://api.authkey.io/request?authkey=${
        process.env.AUTHKEY_API_KEY
      }&mobile=${phone}&country_code=91&sid=${
        process.env.AUTHKEY_SENDER_ID
      }&company=Acediva&otp=${phoneOtp}&message=${encodeURIComponent(
        otpMessage
      )}`
    );

    if (optResponse?.status == "200") {
      return sendResponse(res, 200, "Success", {
        message: "OTP send successfully",
        data: user,
        statusCode: 200,
      });
    } else {
      return sendResponse(res, 422, "Failed", {
        message: "Unable to send OTP",
        statusCode: 200,
      });
    }
  } catch (error) {
    console.error("Error in /send-otp:", error.message);
    // Respond with failure
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error.",
    });
  }
});

userController.post(
  "/sign-up",
  upload.fields([{ name: "profilePic", maxCount: 1 }]),
  async (req, res) => {
    try {
      // Check if the phone number is unique
      const user = await User.findOne({ phone: req.body.phone });
      if (user) {
        return sendResponse(res, 400, "Failed", {
          message: "User is already registered.",
          statusCode: 400,
        });
      }

      // Generate OTP
      const otp = generateOTP();

      // Upload images to Cloudinary
      let profilePic;

      if (req.files["profilePic"]) {
        let image = await cloudinary.uploader.upload(
          req.files["profilePic"][0].path
        );
        profilePic = image.url;
      }

      // Create a new user with provided details
      let newUser = await User.create({
        ...req.body,
        phoneOtp: otp,
        profilePic,
      });

      // Generate JWT token
      const token = jwt.sign(
        { userId: newUser._id, phone: newUser.phone },
        process.env.JWT_KEY
      );

      // Store the token in the user object or return it in the response
      newUser.token = token;
      const updatedUser = await User.findByIdAndUpdate(
        newUser._id,
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
          data: updatedUser,
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

userController.post("/otp-verification", async (req, res) => {
  try {
    const { phone, phoneOtp, firstName } = req.body;
    const user = await User.findOne({ phone, phoneOtp });
    if (user) {
      const updatedUser = await User.findByIdAndUpdate(
        user._id,
        { isPhoneVerified: true, ...(firstName && { firstName }) },
        { new: true }
      );
      return sendResponse(res, 200, "Success", {
        message: "Otp verified successfully",
        data: updatedUser,
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

userController.post("/login", async (req, res) => {
  try {
    const { phone, password } = req.body;
    const user = await User.findOne({ phone, password });
    if (user) {
      return sendResponse(res, 200, "Success", {
        message: "User logged in successfully",
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

userController.post("/resend-otp", async (req, res) => {
  try {
    const { phone } = req.body;
    const user = await User.findOne({ phone });
    if (user) {
      const otp = generateOTP();
      const updatedUser = await User.findByIdAndUpdate(
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
          data: updatedUser,
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

userController.get("/details/:id", auth, async (req, res) => {
  try {
    const id = req.params.id;
    const user = await User.findOne({ _id: id });
    if (user) {
      return sendResponse(res, 200, "Success", {
        message: "User details fetched  successfully",
        data: user,
        statusCode: 200,
      });
    } else {
      return sendResponse(res, 404, "Failed", {
        message: "User not found",
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

userController.post("/list", async (req, res) => {
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
    if (searchKey) query.firstName = { $regex: searchKey, $options: "i" };

    // Construct sorting object
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };

    // Fetch the user list
    const userList = await User.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    // .populate({
    //   path: "product",
    //   select: "name description",
    // })
    // .populate({
    //   path: "createdBy",
    //   select: "name",
    // });
    const totalCount = await User.countDocuments({});
    const activeCount = await User.countDocuments({ status: true });
    sendResponse(res, 200, "Success", {
      message: "User list retrieved successfully!",
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
    });
  }
});

// Helper function to fetch item by type
async function getItemByIdAndType(itemId, itemType) {
  if (!mongoose.Types.ObjectId.isValid(itemId)) return null;
  if (itemType === "Product") return Product.findById(itemId);
  if (itemType === "ComboProduct") return ComboProduct.findById(itemId);
  return null;
}

// Add to Cart
userController.post("/add-to-cart/:id", async (req, res) => {
  try {
    const { id: itemId } = req.params;
    const { userId: currentUserId, itemType } = req.body;

    if (!itemId || !currentUserId || !itemType) {
      return sendResponse(res, 422, "Failed", {
        message: "Missing itemId, userId, or itemType!",
      });
    }

    const item = await getItemByIdAndType(itemId, itemType);
    if (!item) {
      return sendResponse(res, 400, "Failed", {
        message: `${itemType} not found!`,
      });
    }

    const user = await User.findById(currentUserId);
    if (!user) {
      return sendResponse(res, 400, "Failed", { message: "User not found!" });
    }

    if (!Array.isArray(user.cartItems)) user.cartItems = [];

    const cartItemIndex = user.cartItems.findIndex(
      (i) =>
        i.itemId && i.itemId.toString() === itemId && i.itemType === itemType
    );

    let updateQuery, message;

    if (cartItemIndex !== -1) {
      const currentQuantity = user.cartItems[cartItemIndex].quantity;
      if (currentQuantity + 1 > item.stockQuantity) {
        return sendResponse(res, 400, "Failed", {
          message: `Only ${
            item.stockQuantity - currentQuantity
          } item(s) left in stock for this ${itemType}.`,
        });
      }

      updateQuery = {
        $set: {
          [`cartItems.${cartItemIndex}.quantity`]: currentQuantity + 1,
        },
      };
      message = "Item quantity incremented successfully";
    } else {
      if (item.stockQuantity < 1) {
        return sendResponse(res, 400, "Failed", {
          message: `This ${itemType} is currently out of stock.`,
        });
      }

      updateQuery = {
        $push: { cartItems: { itemId, itemType, quantity: 1 } },
      };
      message = `${itemType} added successfully to cart`;
    }

    const updatedUser = await User.findByIdAndUpdate(
      currentUserId,
      updateQuery,
      { new: true }
    );
    sendResponse(res, 200, "Success", {
      message,
      data: updatedUser.cartItems,
      statusCode: 200,
    });
  } catch (error) {
    console.log(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

// Remove from Cart
userController.post("/remove-from-cart/:id", async (req, res) => {
  try {
    const { id: itemId } = req.params;
    const { userId: currentUserId, itemType } = req.body;

    if (!itemId || !currentUserId || !itemType) {
      return sendResponse(res, 422, "Failed", {
        message: "Missing itemId, userId, or itemType!",
      });
    }

    const user = await User.findById(currentUserId);
    if (!user) {
      return sendResponse(res, 400, "Failed", { message: "User not found!" });
    }

    const cartItem = user.cartItems.find(
      (i) =>
        i.itemId && i.itemId.toString() === itemId && i.itemType === itemType
    );

    if (!cartItem) {
      return sendResponse(res, 400, "Failed", { message: "Item not in cart!" });
    }

    let updateQuery, message;

    if (cartItem.quantity > 1) {
      updateQuery = {
        $set: { "cartItems.$[elem].quantity": cartItem.quantity - 1 },
      };
      message = "Item quantity decreased";

      await User.findByIdAndUpdate(currentUserId, updateQuery, {
        new: true,
        arrayFilters: [{ "elem.itemId": itemId, "elem.itemType": itemType }],
      });
    } else {
      updateQuery = {
        $pull: { cartItems: { itemId, itemType } },
      };
      message = "Item removed from cart";

      await User.findByIdAndUpdate(currentUserId, updateQuery, { new: true });
    }

    sendResponse(res, 200, "Success", { message, statusCode: 200 });
  } catch (error) {
    console.log(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

// Remove from Cart (Remove entire item directly)
userController.post("/remove-item-from-cart/:id", async (req, res) => {
  try {
    const { id: itemId } = req.params;
    const { userId: currentUserId, itemType } = req.body;

    if (!itemId || !currentUserId || !itemType) {
      return sendResponse(res, 422, "Failed", {
        message: "Missing itemId, userId, or itemType!",
      });
    }

    const user = await User.findById(currentUserId);
    if (!user) {
      return sendResponse(res, 400, "Failed", { message: "User not found!" });
    }

    const cartItem = user.cartItems.find(
      (i) =>
        i.itemId && i.itemId.toString() === itemId && i.itemType === itemType
    );

    if (!cartItem) {
      return sendResponse(res, 400, "Failed", { message: "Item not in cart!" });
    }

    // Directly remove the entire item from cart
    const updatedUser = await User.findByIdAndUpdate(
      currentUserId,
      {
        $pull: { cartItems: { itemId, itemType } },
      },
      { new: true }
    );

    sendResponse(res, 200, "Success", {
      message: "Item removed completely from cart",
      data: updatedUser.cartItems,
      statusCode: 200,
    });
  } catch (error) {
    console.log(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});


// Get Cart Items
userController.get("/cart/:userId", async (req, res) => {
  try {
    const { userId } = req.params;

    if (!userId) {
      return sendResponse(res, 422, "Failed", {
        message: "User ID is required!",
      });
    }

    const user = await User.findById(userId);
    if (!user) {
      return sendResponse(res, 400, "Failed", { message: "User not found!" });
    }

    let actualTotalAmount = 0;
    let discountedTotalAmount = 0;

    const cartDetails = await Promise.all(
      user.cartItems.map(async (item) => {
        const product = await getItemByIdAndType(item.itemId, item.itemType);
        if (!product) return null;

        const quantity = item.quantity || 1;
        let price = 0;
        let discounted_price = 0;

        if (item.itemType === "Product") {
          price = product.price || 0;
          discounted_price = product.discountedPrice || price;
        } else if (item.itemType === "ComboProduct") {
          price = product.pricing?.actualPrice || 0;
          discounted_price = product.pricing?.comboPrice || price;
        }

        const actualPrice = price * quantity;
        const discountedPrice = discounted_price * quantity;

        actualTotalAmount += actualPrice;
        discountedTotalAmount += discountedPrice;

        return {
          _id: product._id,
          name: product.name,
          itemType: item.itemType,
          productHeroImage: product.productHeroImage || null,
          price,
          discountedPrice: discounted_price,
          quantity,
          totalItemPrice: actualPrice,
          totalItemDiscountedPrice: discountedPrice,
        };
      })
    );

    sendResponse(res, 200, "Success", {
      message: "Cart items retrieved successfully",
      cartItems: cartDetails.filter(Boolean),
      actualTotalAmount,
      discountedTotalAmount,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

userController.post("/add-to-wishlist/:id", async (req, res) => {
  try {
    if (!req.params.id) {
      return sendResponse(res, 422, "Failed", {
        message: "Params not found!",
      });
    }

    const productId = req.params.id;
    const currentUserId = req.body.userId;

    const product = await Product.findOne({ _id: productId });
    if (!product) {
      return sendResponse(res, 400, "Failed", {
        message: "Product not found!",
      });
    }
    const user = await User.findOne({ _id: currentUserId });
    if (!user) {
      return sendResponse(res, 400, "Failed", {
        message: "User not found!",
      });
    }

    let message, updateQuery;
    if (user.wishListItems.includes(productId)) {
      updateQuery = { $pull: { wishListItems: productId } };
      message = "Item removed successfully";
    } else {
      updateQuery = { $push: { wishListItems: productId } };
      message = "Item added successfully";
    }

    // Update the post document with the new array
    await User.findOneAndUpdate({ _id: currentUserId }, updateQuery);

    sendResponse(res, 200, "Success", {
      message: message,
    });
  } catch (error) {
    console.log(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

userController.get("/wishlist/:userId", async (req, res) => {
  try {
    const { userId } = req.params;

    if (!userId) {
      return sendResponse(res, 422, "Failed", {
        message: "User ID is required!",
      });
    }

    const user = await User.findById(userId).populate("wishListItems");

    if (!user) {
      return sendResponse(res, 400, "Failed", {
        message: "User not found!",
      });
    }

    sendResponse(res, 200, "Success", {
      message: "Wishlist items retrieved successfully",
      data: user.wishListItems, // Returns the list of products in the wishlist
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

userController.put("/update", upload.single("profilePic"), async (req, res) => {
  try {
    const id = req.body.userId;
    const userData = await User.findOne({ _id: id });
    if (!userData) {
      return sendResponse(res, 404, "Failed", {
        message: "User not found",
      });
    }

    let updatedData = { ...req.body };

    if (req.file) {
      const profilePic = await cloudinary.uploader.upload(req.file.path);
      updatedData.profilePic = profilePic.url;
    }
    updatedData.profileStatus = "completed";
    const updatedUser = await User.findByIdAndUpdate(id, updatedData, {
      new: true,
    });

    req.io.emit("userUpdated", {
      message: "User profile updated",
      userId: updatedUser._id,
      updatedData: updatedUser,
    });

    sendResponse(res, 200, "Success", {
      message: "User updated successfully!",
      data: updatedUser,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

userController.post("/home-details", async (req, res) => {
  try {
    const homeCategory = await Category.find({});
    const banner = await Banner.find({});
    const trendingProducts = await Product.find({});
    const bestSellerProducts = await Product.find({});
    sendResponse(res, 200, "Success", {
      message: "Home page data fetched successfully!",
      data: {
        homeCategory,
        banner,
        trendingProducts,
        bestSellerProducts,
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

userController.post("/remove-all-from-cart/:id", auth, async (req, res) => {
  try {
    const { id: productId } = req.params;
    const { userId: currentUserId } = req.body;

    if (!productId || !currentUserId) {
      return sendResponse(res, 422, "Failed", {
        message: "Missing productId or userId!",
      });
    }

    const user = await User.findById(currentUserId);
    if (!user) {
      return sendResponse(res, 400, "Failed", { message: "User not found!" });
    }

    const cartItem = user.cartItems.find(
      (item) => item.productId.toString() === productId
    );

    if (!cartItem) {
      return sendResponse(res, 400, "Failed", { message: "Item not in cart!" });
    }

    // ✅ Remove the product entirely from the cart (regardless of quantity)
    await User.findByIdAndUpdate(currentUserId, {
      $pull: { cartItems: { productId } },
    });

    return sendResponse(res, 200, "Success", {
      message: "Product removed from cart",
    });
  } catch (error) {
    console.error(error);
    return sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

userController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const user = await User.findById(id);
    if (!user) {
      return sendResponse(res, 404, "Failed", {
        message: "User not found",
        statusCode: 400,
      });
    }
    await User.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "User deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});
userController.get("/dashboard-details", async (req, res) => {
  try {
    const [
      totalUser,
      activeUser,
      inactiveUser,
      totalBooking,
      activeBooking,
      bookingCompleted,
      totalProduct,
      singleProduct,
      comboProduct,
    ] = await Promise.all([
      User.countDocuments({}),
      User.countDocuments({ profileStatus: "completed" }),
      User.countDocuments({ profileStatus: "incompleted" }),
      Booking.countDocuments({}),
      Booking.countDocuments({
        status: { $in: ["orderPlaced", "orderPacked", "outForDelivery"] },
      }),
      Booking.countDocuments({ status: "completed" }),
      (async () => {
        const [productCount, comboProductCount] = await Promise.all([
          Product.countDocuments({}),
          ComboProduct.countDocuments({}),
        ]);
        return productCount + comboProductCount;
      })(),
      Product.countDocuments({}),
      ComboProduct.countDocuments({}),
    ]);

    // **Last 15 Days Booking Count Logic**
    const last15Days = await Booking.aggregate([
      {
        $match: {
          createdAt: {
            $gte: moment().subtract(15, "days").startOf("day").toDate(),
          },
        },
      },
      {
        $group: {
          _id: { $dateToString: { format: "%Y-%m-%d", date: "$createdAt" } },
          noOfBookings: { $sum: 1 },
        },
      },
      { $sort: { _id: 1 } },
    ]);

    let bookingsLast15Days = [];
    for (let i = 14; i >= 0; i--) {
      let dateObj = moment().subtract(i, "days");
      let formattedDate = dateObj.format("Do MMM"); // "1st Jan" format
      let mongoDate = dateObj.format("YYYY-MM-DD"); // MongoDB ke format ke liye

      let bookingData = last15Days.find((b) => b._id === mongoDate); // Compare in same format

      bookingsLast15Days.push({
        date: formattedDate, // "1st Jan"
        noOfBookings: bookingData ? bookingData.noOfBookings : 0,
        mongoDate: mongoDate,
      });
    }
    sendResponse(res, 200, "Success", {
      message: "Dashboard details retrieved successfully",
      data: {
        users: { totalUser, activeUser, inactiveUser },
        bookings: {
          totalBooking,
          activeBooking,
          bookingCompleted,
        },
        products: {
          totalProduct,
          singleProduct,
          comboProduct,
        },
        last15DaysBookings: bookingsLast15Days, // Reverse for ascending order
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
module.exports = userController;
