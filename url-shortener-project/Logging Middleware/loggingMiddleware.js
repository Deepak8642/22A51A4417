// Logging Middleware/loggingMiddleware.js
const axios = require('axios');

let authToken = null; // Set this after login

function setAuthToken(token) {
  authToken = token;
}

async function Log(stack, level, pkg, message) {
  try {
    if (!authToken) throw new Error("No auth token set for logger");
    await axios.post(
      "http://20.244.56.144/evaluation-service/logs",
      { stack, level, package: pkg, message },
      { headers: { Authorization: `Bearer ${authToken}` } }
    );
  } catch (err) {
    // Do not use console.log, fail silently or handle as needed
  }
}

module.exports = { Log, setAuthToken };
