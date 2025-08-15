const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const productManufactureLocationSchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
  status: {
    type: Boolean,
    default: true,
  }
});

productManufactureLocationSchema.plugin(timestamps);
module.exports = mongoose.model("ProductManufactureLocation", productManufactureLocationSchema);