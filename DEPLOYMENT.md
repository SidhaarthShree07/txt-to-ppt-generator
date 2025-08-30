# Deployment Guide

This document provides instructions for deploying the Text-to-PowerPoint Generator in various environments.

## Local Development

### Quick Start
```bash
# Clone and setup
git clone <https://github.com/SidhaarthShree07/txt-to-ppt-generator>
cd text-to-powerpoint
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
pip install -r requirements.txt

# Run the application
python run.py
# or
python app.py
```

### Development with Debug Mode
```bash
python run.py --debug
```

## Production Deployment

### Using Gunicorn (Recommended)

1. **Install Gunicorn**:
```bash
pip install gunicorn
```

2. **Run with Gunicorn**:
```bash
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

3. **With configuration file** (`gunicorn.conf.py`):
```python
bind = "0.0.0.0:5000"
workers = 4
worker_class = "sync"
worker_connections = 1000
max_requests = 1000
max_requests_jitter = 100
timeout = 30
keepalive = 2
preload_app = True
```

Run with: `gunicorn -c gunicorn.conf.py app:app`

### Docker Deployment

1. **Create Dockerfile**:
```dockerfile
FROM python:3.9-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create non-root user
RUN useradd -m appuser && chown -R appuser:appuser /app
USER appuser

# Expose port
EXPOSE 5000

# Run application
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
```

2. **Build and run**:
```bash
docker build -t text-to-ppt .
docker run -p 5000:5000 text-to-ppt
```

3. **Docker Compose** (`docker-compose.yml`):
```yaml
version: '3.8'
services:
  app:
    build: .
    ports:
      - "5000:5000"
    environment:
      - FLASK_ENV=production
    volumes:
      - ./logs:/app/logs
    restart: unless-stopped
```

## Cloud Deployment

### Heroku

1. **Create `Procfile`**:
```
web: gunicorn -w 4 -b 0.0.0.0:$PORT app:app
```

2. **Deploy**:
```bash
heroku create your-app-name
git push heroku main
```

### Vercel (Recommended for Serverless)

1. **Prerequisites**: The project already includes:
   - `vercel.json` configuration
   - `runtime.txt` specifying Python version
   - Vercel-compatible Flask app structure

2. **Deploy via Vercel CLI**:
```bash
# Install Vercel CLI
npm install -g vercel

# Login to Vercel
vercel login

# Deploy
vercel --prod
```

3. **Deploy via Vercel Dashboard**:
   - Connect your GitHub repository to Vercel
   - Vercel will automatically detect the Flask app
   - No additional configuration needed

4. **Environment Variables**: Set in Vercel dashboard:
   - No environment variables required (users provide their own API keys)

**Note**: Vercel's serverless functions have a 10-second timeout on the Hobby plan and 30 seconds on Pro. For large presentations, consider using other platforms.

### Railway

1. **Create `railway.json`**:
```json
{
  "$schema": "https://railway.app/railway.schema.json",
  "build": {
    "builder": "nixpacks"
  },
  "deploy": {
    "startCommand": "gunicorn -w 4 -b 0.0.0.0:$PORT app:app"
  }
}
```

### Google Cloud Run

1. **Create `.gcloudignore`**:
```
.gcloudignore
.git
.gitignore
README.md
Dockerfile
.dockerignore
node_modules
npm-debug.log
```

2. **Deploy**:
```bash
gcloud run deploy text-to-ppt --source . --platform managed --region us-central1 --allow-unauthenticated
```

### AWS EC2

1. **Install dependencies on EC2**:
```bash
sudo apt update
sudo apt install python3 python3-pip nginx
```

2. **Setup application**:
```bash
git clone <repository-url>
cd text-to-powerpoint
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

3. **Configure Nginx** (`/etc/nginx/sites-available/text-to-ppt`):
```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        client_max_body_size 50M;
    }
}
```

4. **Create systemd service** (`/etc/systemd/system/text-to-ppt.service`):
```ini
[Unit]
Description=Text to PowerPoint Generator
After=network.target

[Service]
User=ubuntu
Group=ubuntu
WorkingDirectory=/home/ubuntu/text-to-powerpoint
Environment=PATH=/home/ubuntu/text-to-powerpoint/venv/bin
ExecStart=/home/ubuntu/text-to-powerpoint/venv/bin/gunicorn -w 4 -b 127.0.0.1:5000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
```

## Environment Variables

Set these environment variables for production:

```bash
export FLASK_ENV=production
export FLASK_DEBUG=False
export SECRET_KEY=your-secret-key-here
export MAX_CONTENT_LENGTH=52428800  # 50MB
```

## Security Considerations

1. **HTTPS**: Always use HTTPS in production
2. **API Keys**: Never log or store API keys
3. **File Upload**: Validate all uploaded files
4. **Rate Limiting**: Implement rate limiting for API endpoints
5. **CORS**: Configure CORS properly if serving from different domains

## Monitoring and Logging

### Basic Logging Setup
```python
import logging
from logging.handlers import RotatingFileHandler

if not app.debug:
    file_handler = RotatingFileHandler('logs/app.log', maxBytes=10240, backupCount=10)
    file_handler.setFormatter(logging.Formatter(
        '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
    ))
    file_handler.setLevel(logging.INFO)
    app.logger.addHandler(file_handler)
    app.logger.setLevel(logging.INFO)
    app.logger.info('Application startup')
```

### Health Check Endpoint
The application includes a health check endpoint at `/api/health` for monitoring.

## Performance Optimization

1. **Caching**: Implement Redis caching for frequent requests
2. **Database**: Add database for user sessions if needed
3. **CDN**: Use CDN for static assets
4. **Load Balancer**: Use load balancer for high traffic

## Troubleshooting

### Common Issues

1. **Module Import Errors**:
   - Ensure all dependencies are installed
   - Check Python path configuration

2. **File Upload Issues**:
   - Verify file size limits
   - Check file permissions

3. **API Key Errors**:
   - Validate API key format
   - Check Gemini API quotas

4. **Memory Issues**:
   - Monitor memory usage during large file processing
   - Implement file cleanup procedures

### Logs Location
- Development: Console output
- Production: `logs/app.log` (if configured)
- Docker: Container logs via `docker logs`

### Debug Mode
Never run with `debug=True` in production. It exposes sensitive information and poses security risks.

## Backup and Recovery

1. **Code**: Use version control (Git)
2. **Config**: Backup configuration files
3. **Logs**: Implement log rotation and archival
4. **Dependencies**: Keep `requirements.txt` updated

For additional support, please check the GitHub issues page.
