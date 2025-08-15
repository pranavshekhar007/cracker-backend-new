const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const roleSchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
  permissions: {
    type: [String],
    required: true,
  },
});
roleSchema.plugin(timestamps);
module.exports = mongoose.model("Role", roleSchema);
