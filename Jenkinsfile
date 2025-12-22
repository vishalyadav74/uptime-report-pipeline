pipeline {
    agent any

    environment {
        PYTHONUNBUFFERED = '1'
        EMAIL_TO = 'vishal.yadav.sys@gmail.com'
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
                  python3 --version
                  python3 -m venv venv
                '''
            }
        }

        stage('Install Dependencies') {
            steps {
                sh '''
                  ./venv/bin/python -m pip install --upgrade pip
                  ./venv/bin/pip install -r requirements.txt
                '''
            }
        }

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  ./venv/bin/python generate_report.py
                '''
            }
        }

        stage('Send HTML Email') {
            steps {
                script {
                    def htmlReport = readFile 'output/uptime_report.html'

                    emailext(
                        subject: "SaaS Application Uptime Report – Weekly & Quarterly",
                        body: htmlReport,
                        mimeType: 'text/html',
                        to: env.EMAIL_TO
                    )
                }
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
            echo '✅ Report generated and HTML email sent successfully'
        }
        failure {
            echo '❌ Pipeline failed'
        }
    }
}
