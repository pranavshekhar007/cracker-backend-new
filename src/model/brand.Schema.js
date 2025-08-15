const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const brandSchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
  status: {
    type: Boolean,
    default: true,
  },
  image: {
    type: String,
    required: true,
  },
  categoryIds:  [{ type: String, ref: "Category" }],
});
brandSchema.plugin(timestamps);
module.exports = mongoose.model("Brand", brandSchema);