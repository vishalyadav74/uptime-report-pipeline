pipeline {
    agent any

    parameters {
        file(name: 'EXCEL_FILE', description: 'Upload Excel file for uptime report')
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
                    def excelPath = ""
                    
                    // Check if file was uploaded
                    if (params.EXCEL_FILE) {
                        excelPath = params.EXCEL_FILE
                        echo "📥 Using uploaded file: ${excelPath}"
                    } else {
                        excelPath = "${WORKSPACE}/uptime_latest1.xlsx"
                        echo "📄 Using default file from repository: ${excelPath}"
                    }
                    
                    // Run Python with file path as argument
                    sh """
                        echo "📊 Generating report..."
                        ./venv/bin/python generate_report.py "${excelPath}"
                    """
                }
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
