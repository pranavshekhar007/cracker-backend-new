const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Product = require("../model/product.Schema");
const ProductFaq = require("../model/productFaq.Schema");
const productFaqController = express.Router();
require("dotenv").config();

productFaqController.post("/create", async (req, res) => {
  try {
    const productFaqCreated = await ProductFaq.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Product Faq added successfully",
      data: productFaqCreated,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

productFaqController.put("/update", async (req, res) => {
  try {
    const id = req.body._id;
    const productFaq = await ProductFaq.findById(id);
    if (!productFaq) {
      return sendResponse(res, 404, "Failed", {
        message: "Product Faq not found",
        statusCode: 403,
      });
    }
    const updatedProductFaq = await ProductFaq.findByIdAndUpdate(id, req.body, {
      new: true,
    });

    sendResponse(res, 200, "Success", {
      message: "Product Faq updated successfully!",
      data: updatedProductFaq,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

// productFaqController.post("/list", async (req, res) => {
//   try {
//     const {
//       searchKey = "",
//       category,
//       pageNo = 1,
//       pageCount = 10,
//       sortByField,
//       sortByOrder,
//     } = req.body;
//     const query = {};
//     if (category) query.category = category;
//     if (searchKey) query.question = { $regex: searchKey, $options: "i" };
//     const sortField = sortByField || "createdAt";
//     const sortOrder = sortByOrder === "asc" ? 1 : -1;
//     const sortOption = { [sortField]: sortOrder };
//     const productFaqList = await ProductFaq.find(query)
//       .sort(sortOption)
//       .limit(parseInt(pageCount))
//       .skip(parseInt(pageNo - 1) * parseInt(pageCount));
//     const productFaq = await ProductFaq.countDocuments({ category: "product" });

//     sendResponse(res, 200, "Success", {
//       message: "Product Faq list retrived successfully.",
//       data: productFaqList,
//       documentCount: { productFaq },
//       statusCode: 200,
//     });
//   } catch (error) {
//     console.error(error);
//     sendResponse(res, 500, "Failed", {
//       message: error.message || "Internal server error",
//       statusCode: 500,
//     });
//   }
// });

productFaqController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      category,
      productId,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;

    const query = {};

    if (category) query.category = category;
    if (searchKey) query.question = { $regex: searchKey, $options: "i" };
    if (productId) query.productId = productId; // Only filter by productId if it exists

    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };

    const productFaqList = await ProductFaq.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip((parseInt(pageNo) - 1) * parseInt(pageCount));

    const totalCount = await ProductFaq.countDocuments(query);

    sendResponse(res, 200, "Success", {
      message: "Product FAQ list retrieved successfully.",
      data: productFaqList,
      documentCount: totalCount,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});


productFaqController.get("/get/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const productFaq = await ProductFaq.findById(id);

    if (!productFaq) {
      return sendResponse(res, 404, "Failed", {
        message: "Product Faq not found",
      });
    }

    sendResponse(res, 200, "Success", {
      message: "Product Faq fetched successfully!",
      data: productFaq,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

productFaqController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const productFaq = await ProductFaq.findById(id);
    if (!productFaq) {
      return sendResponse(res, 404, "Failed", {
        message: "Product Faq not found",
      });
    }
    await ProductFaq.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Product Faq deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 200,
    });
  }
});

module.exports = productFaqController;
