pipeline {
    agent any

    parameters {
        string(
            name: 'MAIL_TO',
            defaultValue: '',
            description: 'Primary recipients (comma separated)'
        )
        string(
            name: 'MAIL_CC',
            defaultValue: '',
            description: 'CC recipients (comma separated)'
        )
    }

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

        stage('Setup Environment') {
            steps {
                sh '''
                  python3 --version
                  python3 -m venv venv
                  ./venv/bin/pip install --upgrade pip
                  ./venv/bin/pip install -r requirements.txt
                '''
            }
        }

        stage('Generate Report & Send Email') {
            steps {
                withCredentials([
                    usernamePassword(
                        credentialsId: 'ITSM_SMTP',
                        usernameVariable: 'SMTP_USER',
                        passwordVariable: 'SMTP_PASSWORD'
                    )
                ]) {
                    withEnv([
                        "MAIL_TO=${params.MAIL_TO}",
                        "MAIL_CC=${params.MAIL_CC}"
                    ]) {
                        sh '''
                          echo "🚀 Generating report & sending email via ITSM SMTP"
                          ./venv/bin/python generate_report.py
                        '''
                    }
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
            echo '✅ Pipeline completed successfully'
        }
        failure {
            echo '❌ Pipeline failed'
        }
    }
}
