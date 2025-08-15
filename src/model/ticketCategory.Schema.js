const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const ticketCategorySchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
  status: {
    type: Boolean,
    default: true,
  }
});

ticketCategorySchema.plugin(timestamps);
module.exports = mongoose.model("TicketCategory", ticketCategorySchema);