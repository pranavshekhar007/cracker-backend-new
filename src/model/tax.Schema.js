const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const taxSchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
  status: {
    type: Boolean,
    default: true,
  },
  percentage:{
    type: Number,
    required : true
  }
});

taxSchema.plugin(timestamps);
module.exports = mongoose.model("Tax", taxSchema);