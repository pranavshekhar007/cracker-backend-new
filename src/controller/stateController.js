const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const State = require("../model/state.schema")

const stateController = express.Router();

stateController.post("/create", async (req, res) => {
  try {
    const { name } = req.body;

    if (!name) {
      return sendResponse(res, 400, "Failed", {
        message: "State name is required",
        statusCode: 400,
      });
    }

    // Ensure stateId increments safely
    const last = await State.findOne({}).sort({ stateId: -1 });
    const lastId = last?.stateId || 0;
    const nextStateId = Number.isInteger(lastId) ? lastId + 1 : 1;

    const stateCreated = await State.create({
      stateId: nextStateId,
      name,
    });

    sendResponse(res, 200, "Success", {
      message: "State created successfully!",
      data: stateCreated,
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


stateController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      status,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;

    const query = {};
    if (status) query.status = status;
    if (searchKey) query.name = { $regex: searchKey, $options: "i" };

    const sortField = sortByField || "stateId";
    const sortOrder = sortByOrder === "desc" ? -1 : 1;
    const sortOption = { [sortField]: sortOrder };

    const stateList = await State.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip((parseInt(pageNo) - 1) * parseInt(pageCount));

    const totalCount = await State.countDocuments({});
    const activeCount = await State.countDocuments({ status: true });

    sendResponse(res, 200, "Success", {
      message: "State list retrieved successfully!",
      data: stateList,
      documentCount: {
        totalCount,
        activeCount,
        inactiveCount: totalCount - activeCount,
      },
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

stateController.put("/update", async (req, res) => {
  try {
    const { _id } = req.body;
    const state = await State.findById(_id);
    if (!state) {
      return sendResponse(res, 404, "Failed", {
        message: "State not found",
        statusCode: 404,
      });
    }

    const updatedState = await State.findByIdAndUpdate(_id, req.body, { new: true });

    sendResponse(res, 200, "Success", {
      message: "State updated successfully!",
      data: updatedState,
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});

stateController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const state = await State.findById(id);
    if (!state) {
      return sendResponse(res, 404, "Failed", {
        message: "State not found",
        statusCode: 404,
      });
    }

    await State.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "State deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});

module.exports = stateController;
