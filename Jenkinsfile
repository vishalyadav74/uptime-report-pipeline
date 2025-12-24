pipeline {
    agent any

    environment {
        PYTHONUNBUFFERED = '1'
        EMAIL_TO = 'yv741518@gmail.com'
    }

    stages {

        stage('Checkout Code') {
            steps {
                git branch: 'main',
                    url: 'https://github.com/vishalyadav74/uptime-report-pipeline.git'
            }
        }

        stage('Setup Virtual Environment') {
            steps {
                sh '''
                  python3 -m venv venv
                '''
            }
        }

        stage('Install Dependencies') {
            steps {
                sh '''
                  ./venv/bin/pip install -r requirements.txt
                '''
            }
        }

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  echo "Using Excel file: $UPTIME_EXCEL"
                  ./venv/bin/python generate_report.py
                '''
            }
        }

        stage('Send HTML Email') {
            steps {
                script {
                    def htmlReport = readFile 'output/uptime_report.html'

                    emailext(
                        subject: "SaaS Application Uptime Report",
                        body: htmlReport,
                        mimeType: 'text/html',
                        to: env.EMAIL_TO
                    )
                }
            }
        }
    }
}
