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

        stage('Prepare Excel File') {
            steps {
                script {
                    if (params.UPLOADED_EXCEL) {
                        // File uploaded - use uploaded file
                        def uploadedFile = params.UPLOADED_EXCEL
                        echo "✅ File uploaded: ${uploadedFile}"
                        
                        // Set environment variable for Python script
                        env.UPTIME_EXCEL = uploadedFile
                    } else {
                        // No file uploaded - use default file
                        def defaultFile = "${WORKSPACE}/uptime_latest1.xlsx"
                        echo "📄 Using default file: ${defaultFile}"
                        
                        // Set environment variable for Python script
                        env.UPTIME_EXCEL = defaultFile
                    }
                }
            }
        }

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  echo "📊 Generating report..."
                  echo "Excel file path: ${UPTIME_EXCEL}"
                  ./venv/bin/python generate_report.py
                '''
            }
        }

        stage('Send HTML Email') {
            steps {
                script {
                    def outputPath = "${WORKSPACE}/output/uptime_report.html"
                    
                    if (fileExists(outputPath)) {
                        def htmlReport = readFile outputPath
                        
                        emailext(
                            subject: "SaaS Application Uptime Report – Weekly & Quarterly",
                            body: htmlReport,
                            mimeType: 'text/html',
                            to: env.EMAIL_TO,
                            from: 'Jenkins <yv741518@gmail.com>',
                            replyTo: 'yv741518@gmail.com'
                        )
                        echo "✅ Email sent successfully"
                    } else {
                        error "❌ Report file not found at: ${outputPath}"
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
