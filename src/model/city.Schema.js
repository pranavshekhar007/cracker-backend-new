const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const citySchema = new mongoose.Schema({
  cityId: {
    type: Number,
    unique: true,
    index: true,
  },
  name: {
    type: String,
    required: true,
    trim: true,
  },
  stateId: {
    type: Number,
    required: true,
  },
  minimumPrice: {
    type: Number,
    required: true,
    min: 0,
  },
  deliveryCharge: {
    type: Number,
    required: true,
    min: 0,
  },
  status: {
    type: Boolean,
    default: true,
  },
});

citySchema.plugin(timestamps);
module.exports = mongoose.model("City", citySchema);
