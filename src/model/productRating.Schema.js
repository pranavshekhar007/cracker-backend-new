const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const productRatingSchema = mongoose.Schema({
    review: {
        type: String,
        required: true,
    },
    rating: {
        type: Number,
        default: true,
    },
    productId: {
        type: mongoose.Schema.Types.ObjectId,
        ref: "Product",
        required: true, 
    },
    userId: {
        type: mongoose.Schema.Types.ObjectId,
        ref: "User",
        required: true, 
    },
});

productRatingSchema.plugin(timestamps);
module.exports = mongoose.model("ProductRating", productRatingSchema);