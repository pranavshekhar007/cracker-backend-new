const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const productTypeSchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
  status: {
    type: Boolean,
    default: true,
  }
});

productTypeSchema.plugin(timestamps);
module.exports = mongoose.model("ProductType", productTypeSchema);