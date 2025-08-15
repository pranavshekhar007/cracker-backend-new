const express = require("express");
const Pincode = require("../model/pincode.Schema");
const { sendResponse } = require("../utils/common");
require("dotenv").config();

const pincodeController = express.Router();

pincodeController.post("/create", async (req, res) => {
  try {
    const { pincode, cityId } = req.body;

    if (!pincode || !cityId) {
      return sendResponse(res, 400, "Failed", {
        message: "Pincode and cityId are required",
        statusCode: 400,
      });
    }

    const last = await Pincode.findOne({}).sort({ pincodeId: -1 });
    const lastId = last?.pincodeId || 0;
    const nextPincodeId = Number.isInteger(lastId) ? lastId + 1 : 1;

    const created = await Pincode.create({
      pincodeId: nextPincodeId,
      pincode,
      cityId,
    });

    sendResponse(res, 200, "Success", {
      message: "Pincode created successfully!",
      data: created,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});

pincodeController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      status,
      cityId,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;

    const query = {};
    if (status) query.status = status;
    if (searchKey) query.pincode = { $regex: searchKey, $options: "i" };
    if (cityId) query.cityId = cityId;

    const sortField = sortByField || "pincodeId";
    const sortOrder = sortByOrder === "desc" ? -1 : 1;

    const pincodeList = await Pincode.find(query)
      .sort({ [sortField]: sortOrder })
      .limit(parseInt(pageCount))
      .skip((parseInt(pageNo) - 1) * parseInt(pageCount));

    const totalCount = await Pincode.countDocuments({});
    const activeCount = await Pincode.countDocuments({ status: true });

    sendResponse(res, 200, "Success", {
      message: "Pincode list retrieved successfully!",
      data: pincodeList,
      documentCount: {
        totalCount,
        activeCount,
        inactiveCount: totalCount - activeCount,
      },
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



pincodeController.put("/update", async (req, res) => {
  try {
    const { _id, pincode } = req.body;

    if (!_id || !pincode) {
      return sendResponse(res, 400, "Failed", {
        message: "_id and pincode are required",
        statusCode: 400,
      });
    }

    const updated = await Pincode.findByIdAndUpdate(
      _id,
      { pincode },
      { new: true }
    );

    if (!updated) {
      return sendResponse(res, 404, "Failed", {
        message: "Pincode not found",
        statusCode: 404,
      });
    }

    sendResponse(res, 200, "Success", {
      message: "Pincode updated successfully!",
      data: updated,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});


pincodeController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;

    const deleted = await Pincode.findByIdAndDelete(id);

    if (!deleted) {
      return sendResponse(res, 404, "Failed", {
        message: "Pincode not found",
        statusCode: 404,
      });
    }

    sendResponse(res, 200, "Success", {
      message: "Pincode deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});

// Get pincodes by cityId
// pincodeController.post("/get-by-city", async (req, res) => {
//   try {
//     const { cityId } = req.body;

//     if (!cityId) {
//       return sendResponse(res, 400, "Failed", {
//         message: "cityId is required",
//         statusCode: 400,
//       });
//     }

//     const pincodes = await Pincode.find({ cityId });

//     sendResponse(res, 200, "Success", {
//       message: "Pincode list fetched by cityId successfully!",
//       data: pincodes,
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


module.exports = pincodeController;
