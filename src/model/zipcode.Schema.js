const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const zipcodeSchema = mongoose.Schema({
  zipcode: {
    type: String,
    required: true,
  },
  city: {
    type: String,
    required: true,
  },
  minFreeDeliveryOrderAmount: {
    type: String,
    required: true,
  },
  deliveryCharges: {
    type: String,
    required: true,
  }
});

zipcodeSchema.plugin(timestamps);
module.exports = mongoose.model("Zipcode", zipcodeSchema);