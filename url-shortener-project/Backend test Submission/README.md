# URL Shortener Microservice

This project implements a robust HTTP URL Shortener Microservice that provides core URL shortening functionality along with basic analytical capabilities for the shortened links.

## Table of Contents

- [Technologies Used](#technologies-used)
- [Setup Instructions](#setup-instructions)
- [API Endpoints](#api-endpoints)
- [Error Handling](#error-handling)
- [Logging](#logging)

## Technologies Used

- Node.js
- Express.js
- MongoDB (or any other database of your choice)
- Mongoose (for MongoDB object modeling)
- Custom Logging Middleware

## Setup Instructions

1. Clone the repository:
   ```
   git clone <repository-url>
   ```

2. Navigate to the backend directory:
   ```
   cd url-shortener-project/backend
   ```

3. Install the dependencies:
   ```
   npm install
   ```

4. Start the server:
   ```
   npm start
   ```

5. The server will run on `http://localhost:port`, where `port` is defined in your app configuration.

## API Endpoints

### Create Short URL

- **Method:** POST
- **Route:** `/shorturls`
- **Request Body:**
  ```json
  {
    "url": "https://example.com/very-long-url",
    "validity": 30,
    "shortcode": "abcd1"
  }
  ```
- **Response:**
  ```json
  {
    "shortLink": "https://hostname:port/abcd1",
    "expiry": "2025-01-01T00:30:00Z"
  }
  ```

### Retrieve Short URL Statistics

- **Method:** GET
- **Route:** `/shorturls/:shortcode`
- **Response:**
  ```json
  {
    "originalUrl": "https://example.com/very-long-url",
    "creationDate": "2025-01-01T00:00:00Z",
    "expiryDate": "2025-01-01T00:30:00Z",
    "clickCount": 10,
    "clickData": [
      {
        "timestamp": "2025-01-01T00:05:00Z",
        "referrer": "https://google.com",
        "location": "New York, USA"
      }
    ]
  }
  ```

## Error Handling

The API will return appropriate HTTP status codes and descriptive JSON response bodies for invalid requests, such as:

- 400 Bad Request for malformed input
- 404 Not Found for non-existent shortcodes
- 409 Conflict for shortcode collisions
- 410 Gone for expired links

## Logging

The application uses a custom logging middleware to log all requests and responses, ensuring that all interactions are recorded for monitoring and debugging purposes.
