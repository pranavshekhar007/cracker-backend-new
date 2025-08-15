const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const notifySchema = mongoose.Schema({
  icon: {
    type: String,
  },
  title: {
    type: String,
  },
  subTitle: {
    type: String,
  },
  notifyUserIds: {
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
});

notifySchema.plugin(timestamps);
module.exports = mongoose.model("Notify", notifySchema);
