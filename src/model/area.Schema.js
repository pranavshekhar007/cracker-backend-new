const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const areaSchema = new mongoose.Schema({
  areaId: {
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
  cityId: {
    type: Number,
    required: true,
  },
  pincodeId: {
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

areaSchema.plugin(timestamps);
module.exports = mongoose.model("Area", areaSchema);
