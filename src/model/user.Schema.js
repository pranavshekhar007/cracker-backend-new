const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const userSchema = mongoose.Schema({
  firstName: {
    type: String,
  },
  lastName: {
    type: String,
  },
  subscriptionDetails: {
    subscriptionId: {
      type: mongoose.Schema.Types.ObjectId,
      ref: "SubscriptionChit",
    },
    planName: { type: String },
    planPrice: { type: String },
    duration: { type: Number },
    startDate: { type: Date },
    endDate: { type: Date },
  },
  email: {
    type: String,
  },
  password: {
    type: String,
  },
  emailOtp: {
    type: String,
  },
  phoneOtp: {
    type: String,
  },
  token: {
    type: String,
  },
  phone: {
    type: String,
    required: true,
  },
  profilePic: {
    type: String,
  },
  isEmailVerified: {
    type: Boolean,
    default: false,
  },
  isPhoneVerified: {
    type: Boolean,
    default: false,
  },
  countryCode: {
    type: String,
    default: "91",
  },
  profileStatus:{
    type: String,
    default: "incompleted",
      enum: ["incompleted", "completed"],
  },
  pincode: {
    type: String,
  },
  address: {
    type: String,
  },
  cartItems: [{
    itemId: { type: String, required: true },
    itemType: { type: String, enum: ["Product", "ComboProduct"], required: true },
    quantity: { type: Number },
  }],
    wishListItems: [{ type: String, ref: "Product" }],
});

userSchema.plugin(timestamps);
module.exports = mongoose.model("User", userSchema);
