const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const supportSchema = mongoose.Schema({
  userPrivacyPolicy: {
    type: String,
  },
  userTermsAndCondition: {
    type: String,
  },
  userCookiePolicy: {
    type: String,
  },
  userShippingPolicy: {
    type: String,
  },
  refundAndReturn: {
    type: String,
  },
  supportContact: {
    type: String,
  },
  supportEmail: {
    type: String,
  },


});

supportSchema.plugin(timestamps);
module.exports = mongoose.model("Support", supportSchema);
