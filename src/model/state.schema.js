const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const stateSchema = mongoose.Schema({
  stateId: {
    type: Number,
    unique: true,
    index: true,
  },
  name: {
    type: String,
    required: true,
  },

  status: {
    type: Boolean,
    default: true,
  },
});

stateSchema.plugin(timestamps);
module.exports = mongoose.model("State", stateSchema);
