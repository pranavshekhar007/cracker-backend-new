const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const permissionSchema = mongoose.Schema({
  name: {
    type: String,
    required: true,
  },
});
permissionSchema.plugin(timestamps);
module.exports = mongoose.model("Permission", permissionSchema);
