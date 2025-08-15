const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const subCategorySchema = mongoose.Schema({
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
    },
    categoryId: {
        type: mongoose.Schema.Types.ObjectId,
        ref: "Category",
        required: true, 
    },
    specialApperence: {
        type: String,
      },
});

subCategorySchema.plugin(timestamps);
module.exports = mongoose.model("SubCategory", subCategorySchema);