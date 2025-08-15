const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const ticketSchema = mongoose.Schema({
  subject: {
    type: String,
    required: true,
  },
  userId: {
    type: String,
    required: true,
  },
  image: {
    type: String,
    // required: true,
  },
  userType: {
    type: String,
    enum: ["User", "Vender", "Driver"],
  },
  status: {
    type: Boolean,
    default: true,
  },
  ticketCategoryId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "TicketCategory",
    required: true,
  },
});

ticketSchema.plugin(timestamps);
module.exports = mongoose.model("Ticket", ticketSchema);
