FROM node:20.18.1-alpine AS builder

WORKDIR /app

# Install build dependencies
RUN apk add --no-cache python3 make g++

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy the rest of the application
COPY . .

# Set NODE_ENV to production
ENV NODE_ENV=production

# Build the application
RUN npm run build

# Use nginx to serve the files
FROM nginx:alpine

# Copy the built files from builder stage
COPY --from=builder /app/dist /usr/share/nginx/html
# Explicitly copy assets directory
COPY --from=builder /app/assets /usr/share/nginx/html/assets

# Copy nginx configuration
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Expose port 80
EXPOSE 80

# Start nginx
CMD ["nginx", "-g", "daemon off;"] 