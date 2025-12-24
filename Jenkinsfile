pipeline {
    agent any

    parameters {
        file(name: 'UPLOADED_EXCEL', description: 'Upload Excel file for uptime report')
    }

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
                script {
                    echo "🔍 Checking for uploaded file..."
                    
                    def excelPath = ""
                    
                    if (params.UPLOADED_EXCEL) {
                        echo "✅✅✅ FILE UPLOADED DETECTED! ✅✅✅"
                        echo "Uploaded file: ${params.UPLOADED_EXCEL}"
                        excelPath = params.UPLOADED_EXCEL
                    } else {
                        echo "⚠️⚠️⚠️ NO FILE UPLOADED ⚠️⚠️⚠️"
                        echo "Using default file from repository"
                        // ✅ CORRECTED FILE NAME - use uptime_latest.xlsx instead of uptime_latest1.xlsx
                        excelPath = "${WORKSPACE}/uptime_latest.xlsx"
                    }
                    
                    echo "📊 Final file path: ${excelPath}"
                    
                    // Run Python with file path
                    sh """
                        ./venv/bin/python generate_report.py "${excelPath}"
                    """
                }
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
                        to: env.EMAIL_TO,
                        from: 'Jenkins <yv741518@gmail.com>',
                        replyTo: 'yv741518@gmail.com'
                    )
                    echo "✅ Email sent successfully"
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
