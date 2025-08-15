const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const tagSchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
  slug: {
    type: String,
    required:true,
  },
  description: {
    type: String,
    default: "",
  },
  status: {
    type: Boolean,
    default: true,
  }
});

tagSchema.plugin(timestamps);
module.exports = mongoose.model("Tag", tagSchema);