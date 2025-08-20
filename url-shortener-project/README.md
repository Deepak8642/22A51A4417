# URL Shortener Project

This project is a robust HTTP URL Shortener Microservice that provides core URL shortening functionality along with basic analytical capabilities for the shortened links. It consists of a backend microservice built with Node.js and Express, and a frontend web application developed using React.

## Project Structure

```
url-shortener-project
├── backend
│   ├── src
│   │   ├── app.js
│   │   ├── routes
│   │   │   └── shorturls.js
│   │   ├── middleware
│   │   │   └── loggingMiddleware.js
│   │   ├── controllers
│   │   │   └── shorturlController.js
│   │   ├── models
│   │   │   └── shorturl.js
│   │   └── utils
│   │       └── shortcodeGenerator.js
│   ├── package.json
│   └── README.md
├── frontend
│   ├── public
│   │   └── index.html
│   ├── src
│   │   ├── App.js
│   │   ├── index.js
│   │   ├── components
│   │   │   ├── ShortenerForm.js
│   │   │   ├── ShortenedLinksList.js
│   │   │   └── StatisticsPage.js
│   │   └── utils
│   │       └── api.js
│   ├── package.json
│   └── README.md
└── README.md
```

## Backend

The backend microservice is responsible for handling URL shortening requests and providing analytics for the shortened links. It includes the following key components:

- **app.js**: Entry point of the backend application, setting up the Express server and middleware.
- **routes/shorturls.js**: Defines the API routes for creating short URLs and retrieving statistics.
- **middleware/loggingMiddleware.js**: Implements logging for all requests and responses.
- **controllers/shorturlController.js**: Contains the logic for creating short URLs and retrieving statistics.
- **models/shorturl.js**: Defines the data model for storing short URLs and their metadata.
- **utils/shortcodeGenerator.js**: Utility functions for generating unique short codes.

### Setup Instructions

1. Navigate to the `backend` directory.
2. Install dependencies using `npm install`.
3. Start the server with `npm start`.

## Frontend

The frontend application provides a user interface for interacting with the URL shortener service. It includes:

- **App.js**: Main component that sets up routing and integrates various components.
- **components/ShortenerForm.js**: Allows users to input URLs and optional parameters for shortening.
- **components/ShortenedLinksList.js**: Displays the list of shortened URLs and their expiry dates.
- **components/StatisticsPage.js**: Retrieves and displays statistics for the shortened URLs.

### Setup Instructions

1. Navigate to the `frontend` directory.
2. Install dependencies using `npm install`.
3. Start the application with `npm start`.

## Usage

- Use the URL Shortener page to input long URLs, specify validity periods, and optionally provide custom shortcodes.
- View the shortened links along with their expiry dates after successful creation.
- Access the Statistics page to view click data and other analytics for the shortened URLs.

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue for any enhancements or bug fixes.

## License

This project is licensed under the MIT License.