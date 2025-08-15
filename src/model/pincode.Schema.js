const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const pincodeSchema = new mongoose.Schema({
  pincodeId: {
    type: Number,
    unique: true,
    index: true,
  },
  cityId: {
    type: Number,
    required: true,
  },

  pincode: {
    type: String,
    required: true,
    trim: true,
  },
  status: {
    type: Boolean,
    default: true,
  },
});

pincodeSchema.plugin(timestamps);
module.exports = mongoose.model("Pincode", pincodeSchema);
