const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const categorySchema = mongoose.Schema({
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
  specialApperence: {
    type: String,
  },
});

categorySchema.plugin(timestamps);
module.exports = mongoose.model("Category", categorySchema);