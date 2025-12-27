pipeline {
    agent any

    parameters {
        string(
            name: 'MAIL_TO',
            defaultValue: 'incident@businessnext.com',
            description: 'Primary recipient'
        )

        string(
            name: 'MAIL_CC',
            defaultValue: 'itsm@businessnext.com,ops@businessnext.com',
            description: 'CC list (comma separated)'
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

        stage('Generate Report') {
            steps {
                script {
                    echo "🚀 Generating uptime report..."

                    if (params.UPLOADED_EXCEL) {
                        sh "./venv/bin/python generate_report.py '${params.UPLOADED_EXCEL}'"
                    } else {
                        sh "./venv/bin/python generate_report.py"
                    }
                }
            }
        }

        stage('Send Email') {
            steps {
                script {
                    def htmlReport = readFile 'output/uptime_report.html'

                    emailext(
                        subject: "SAAS Accounts Weekly & Quarterly Application Uptime Report",
                        body: htmlReport,
                        mimeType: 'text/html',

                        to: params.MAIL_TO,
                        cc: params.MAIL_CC,

                        from: 'incident@businessnext.com',
                        replyTo: 'incident@businessnext.com'
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
            echo '✅ Pipeline completed successfully'
        }
        failure {
            echo '❌ Pipeline failed'
        }
    }
}
