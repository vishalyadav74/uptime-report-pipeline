pipeline {
    agent any

    parameters {
        file(name: 'UPLOADED_EXCEL', description: 'Upload Excel file for uptime report')
    }

    environment {
        PYTHONUNBUFFERED = '1'
        EMAIL_TO = 'yv741518@gmail.com'
        UPTIME_EXCEL = ''
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
                    // Check if file was uploaded via parameter
                    if (params.UPLOADED_EXCEL) {
                        echo "📥 File uploaded via parameter detected!"
                        
                        // Get the uploaded file path
                        def uploadedFile = params.UPLOADED_EXCEL
                        
                        // Create input directory if it doesn't exist
                        sh 'mkdir -p input'
                        
                        // Copy uploaded file to workspace with proper name
                        sh """
                          cp "${uploadedFile}" input/uptime.xlsx
                        """
                        
                        env.UPTIME_EXCEL = "${WORKSPACE}/input/uptime.xlsx"
                        echo "✅ Using UPLOADED file: ${UPTIME_EXCEL}"
                        
                    } else {
                        // Use the default file from repository
                        env.UPTIME_EXCEL = "${WORKSPACE}/uptime_latest1.xlsx"
                        echo "📄 Using DEFAULT file from repository: ${UPTIME_EXCEL}"
                    }
                }
            }
        }

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  echo "📊 Generating report from $UPTIME_EXCEL"
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
