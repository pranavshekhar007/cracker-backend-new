const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const schemeConfigSchema = mongoose.Schema({
  schemeStartDate: {
    type: Date,
    required: true,
  },
  schemeEndDate: {
    type: Date,
    required: true,
  },
});

schemeConfigSchema.plugin(timestamps);

module.exports = mongoose.model("SchemeConfig", schemeConfigSchema);
