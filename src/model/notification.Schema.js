const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const notificationSchema = mongoose.Schema({
  icon: {
    type: String,
  },
  title: {
    type: String,
  },
  subTitle: {
    type: String,
  },
  notifyUserId: {
    type: String
  },
  category: {
    type: String,
  },
  subCategory: {
    type: String,
  },
  isRead: {
    type: Boolean,
    default: false,
  },
  notifyUser:{
    type: String, enum: ["User", "Vender", "Driver", "Admin"],
  },
});

notificationSchema.plugin(timestamps);
module.exports = mongoose.model("Notification", notificationSchema);
