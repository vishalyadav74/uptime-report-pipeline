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

        stage('Generate Report') {
            steps {
                script {
                    echo "🚀 Generating uptime report..."
                    sh "./venv/bin/python generate_report.py"
                }
            }
        }

        stage('Send Email') {
            steps {
                script {
                    def htmlReport = readFile 'output/uptime_report.html'

                    // ✅ FIX: combine TO + CC safely
                    def recipients = params.MAIL_TO
                    if (params.MAIL_CC?.trim()) {
                        recipients = recipients ?
                            "${params.MAIL_TO},${params.MAIL_CC}" :
                            params.MAIL_CC
                    }

                    emailext(
                        subject: "SAAS Accounts Weekly & Quarterly Application Uptime Report",
                        body: htmlReport,
                        mimeType: 'text/html',

                        // 🔥 ONLY use `to`
                        to: recipients,

                        from: 'incident@businessnext.com',
                        replyTo: 'incident@businessnext.com'
                    )

                    echo "✅ Email sent to: ${recipients}"
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
