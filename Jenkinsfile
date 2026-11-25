pipeline {
  agent any
  stages {
    stage('Checkout') { steps { checkout scm } }
    stage('Backend') {
      steps {
        sh 'python3 -m pip install -r backend/requirements.txt pytest pytest-cov'
        sh 'cd backend && pytest --cov=. --cov-report=xml'
      }
    }
    stage('Frontend') {
      steps { sh 'cd frontend && npm ci && npm run build' }
    }
    stage('Package') {
      steps { sh 'docker compose build' }
    }
  }
}
