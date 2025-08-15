const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const bookingSchema = mongoose.Schema({
  totalAmount: {
    type: Number,
  },
  status: {
    type: String,
    default: "pending",
    enum: [
      "pending",
      "paymentSsUpload",
      "approved",
      "ssRejected",
      "cancelled",
      "orderPlaced",
      "orderPacked",
      "shipping",
      "outForDelivery",
      "completed",
    ],
  },
  signature: {
    type: String,
    // require: true,
  },

  modeOfPayment: {
    type: String,
    enum: ["COD", "Online"],
  },
  paymentId: {
    type: String,
  },
  paymentSs: {
    type: String,
  },
  product: [
    {
      productId: { type: String, ref: "Product" },
      quantity: { type: Number },
      totalPrice: { type: Number },
      productHeroImage: { type: String },
    },
  ],
  comboProduct: [
    {
      comboProductId: { type: String, ref: "ComboProduct" },
      quantity: { type: Number },
      totalPrice: { type: Number },
      productHeroImage: { type: String },
    },
  ],


  shipping: {
    type: String,
    enum: ["homeDelivery", "lorryPay"],
  },

  deliveryCharge: {
    type: String,
  },

  userId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "User",
  },
  cityId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "City",
  },

  address: {
    fullName: { type: String, required: true },
    phone: { type: String, required: true },
    alternatePhone: { type: String },
    landmark: { type: String },
    area: { type: String },
    city: { type: String, required: true },
    state: { type: String, required: true },
    pincode: { type: String, required: true },
    country: { type: String, required: true },
  },

  statusHistory: [
    {
      status: { type: String },
      updatedAt: { type: Date, default: Date.now },
    },
  ],

  //  couponId: {
  //     type: mongoose.Schema.Types.ObjectId,
  //     ref: "Coupon",
  //   },
  //   couponDiscountValue: {
  //     type: Number,
  //   },
});

bookingSchema.plugin(timestamps);
module.exports = mongoose.model("Booking", bookingSchema);
