import multiprocessing

# Gunicorn configuration
bind = "0.0.0.0:8000"
workers = 1  # Keep it minimal for Render's free tier
worker_class = "gthread"  # Use threads instead of processes
threads = 2  # Number of threads per worker
timeout = 300  # Increase timeout to 5 minutes
keepalive = 24
worker_connections = 1000
max_requests = 100
max_requests_jitter = 50

# Logging
accesslog = "-"
errorlog = "-"
loglevel = "info"

# SSL
forwarded_allow_ips = "*"
secure_scheme_headers = {"X-Forwarded-Proto": "https"} 