const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const Support = require("../model/support.Schema");
const Contact = require("../model/contact.Schema");
const Faq = require("../model/faq.Schema");
const supportController = express.Router();
require("dotenv").config();

supportController.post("/add-contact-query", async (req, res) => {
  try {
    const contactCreated = await Contact.create(req.body);
    sendResponse(res, 200, "Success", {
      message:
        "We have recived your query, Someone from our team will reach back to you!",
      data: contactCreated,
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
supportController.post("/list-contact-query", async (req, res) => {
  try {
    const {
      searchKey = "",
      category,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;
    const query = {};
    if (category) query.category = category;
    if (searchKey) query.question = { $regex: searchKey, $options: "i" };
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };
    const contactList = await Contact.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const userQueries = await Contact.countDocuments({ category: "user" });
    const driverQueries = await Contact.countDocuments({ category: "driver" });
    const vendorQueries = await Contact.countDocuments({ category: "vender" });

    sendResponse(res, 200, "Success", {
      message: "Contact list retrived successfully.",
      data: contactList,
      documentCount:{userQueries, driverQueries, vendorQueries},
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
supportController.get("/details", async (req, res) => {
  try {
    const supportDetails = await Support.findOne({});
    sendResponse(res, 200, "Success", {
      message: "Support details retrived successfully",
      data: supportDetails,
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
supportController.post("/add-details", async (req, res) => {
  try {
    const supportDetails = await Support.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Support created successfully",
      data: supportDetails,
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
supportController.put("/update-details", async (req, res) => {
  try {
    const id = req.body._id;
    const support = await Support.findById(id);
    if (!support) {
      return sendResponse(res, 404, "Failed", {
        message: "Support not found",
        statusCode: 403,
      });
    }
    const updateSupport = await Support.findByIdAndUpdate(id, req.body, {
      new: true,
    });

    sendResponse(res, 200, "Success", {
      message: "Support updated successfully!",
      data: updateSupport,
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
supportController.get("/privacy-policy", async (req, res) => {
  try {
    const supportDetails = await Support.findOne({});
    res.setHeader("Content-Type", "text/html");
    res.send(supportDetails.privacyPolicy);
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

supportController.post("/create-faq", async (req, res) => {
  try {
    const faqCreated = await Faq.create(req.body);
    sendResponse(res, 200, "Success", {
      message: "Faq added successfully",
      data: faqCreated,
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
supportController.put("/update-faq", async (req, res) => {
  try {
    const id = req.body._id;
    const faq = await Faq.findById(id);
    if (!faq) {
      return sendResponse(res, 404, "Failed", {
        message: "Faq not found",
        statusCode: 403,
      });
    }
    const updatedFaq = await Faq.findByIdAndUpdate(id, req.body, {
      new: true,
    });

    sendResponse(res, 200, "Success", {
      message: "Faq updated successfully!",
      data: updatedFaq,
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
supportController.post("/list-faq", async (req, res) => {
  try {
    const {
      searchKey = "",
      category,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;
    const query = {};
    if (category) query.category = category;
    if (searchKey) query.question = { $regex: searchKey, $options: "i" };
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };
    const faqList = await Faq.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const userFaq = await Faq.countDocuments({ category: "user" });
    const driverFaq = await Faq.countDocuments({ category: "driver" });
    const vendorFaq = await Faq.countDocuments({ category: "vendor" });

    sendResponse(res, 200, "Success", {
      message: "Faq list retrived successfully.",
      data: faqList,
      documentCount:{userFaq, driverFaq, vendorFaq},
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
supportController.delete("/delete-faq/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const faq = await Faq.findById(id);
    if (!faq) {
      return sendResponse(res, 404, "Failed", {
        message: "Faq not found",
      });
    }
    await Faq.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Faq deleted successfully!",
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

module.exports = supportController;
