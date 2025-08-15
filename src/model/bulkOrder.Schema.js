const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const bulkOrderSchema = mongoose.Schema({
  firstName: {
    type: String,
  },
  lastName: {
    type: String,
  },
  contactNumber: {
    type: Number,
  },
  message: {
    type: String,
  },
  image: {
    type: String,
  },
  status: {
    type: String,
    default: "orderPlaced",
    enum: ["orderPlaced", "completed", "cancelled"],
  },
});

bulkOrderSchema.plugin(timestamps);
module.exports = mongoose.model("BulkOrder", bulkOrderSchema);
