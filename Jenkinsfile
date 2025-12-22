pipeline {
    agent any

    environment {
        PYTHONUNBUFFERED = '1'
    }

    stages {

        stage('Checkout Code') {
            steps {
                git branch: 'main',
                    url: 'https://github.com/vishalyadav74/uptime-report-pipeline.git'
            }
        }

        stage('Install Dependencies') {
            steps {
                sh '''
                  python3 --version
                  python3 -m pip install --upgrade pip
                  python3 -m pip install -r requirements.txt
                '''
            }
        }

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  python3 generate_report.py
                '''
            }
        }

        stage('Archive Report') {
            steps {
                archiveArtifacts artifacts: 'output/uptime_report.html', fingerprint: true
            }
        }
    }

    post {
        success {
            echo '✅ Uptime report generated successfully'
        }
        failure {
            echo '❌ Pipeline failed'
        }
    }
}
