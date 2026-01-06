pipeline {
    agent any

    parameters {
        string(name: 'MAIL_TO', defaultValue: '', description: 'To (comma separated)')
        string(name: 'MAIL_CC', defaultValue: '', description: 'CC (comma separated)')
    }

    environment {
        PYTHONUNBUFFERED = '1'
    }

    stages {

        stage('Checkout') {
            steps {
                git branch: 'main',
                    url: 'https://github.com/vishalyadav74/uptime-report-pipeline.git'
            }
        }

        stage('Setup') {
            steps {
                sh '''
                  python3 --version
                  python3 -m venv venv
                  ./venv/bin/pip install --upgrade pip
                  ./venv/bin/pip install -r requirements.txt
                '''
            }
        }

        stage('Generate Report') {
            steps {
                sh "./venv/bin/python generate_report.py"
            }
        }

        stage('Send Email') {
            steps {
                sh """
                  ./venv/bin/python send.py \
                    --subject "SAAS Accounts Weekly & Quarterly Application Uptime Report" \
                    --to "${params.MAIL_TO}" \
                    --cc "${params.MAIL_CC}"
                """
            }
        }

        stage('Archive') {
            steps {
                archiveArtifacts artifacts: 'output/*', fingerprint: true
            }
        }
    }
}
