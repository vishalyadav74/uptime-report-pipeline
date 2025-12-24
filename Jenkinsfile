pipeline {
    agent any

    parameters {
        file(name: 'UPLOADED_EXCEL', description: 'Optional: Upload different Excel file')
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

        stage('Setup Environment') {
            steps {
                sh '''
                  python3 --version
                  python3 -m venv venv
                  ./venv/bin/pip install --upgrade pip
                  ./venv/bin/pip install -r requirements.txt
                  
                  echo "📁 Repository contents:"
                  ls -la
                  echo ""
                  echo "🔍 Excel files found:"
                  find . -name "*.xlsx" -o -name "*.xls" | head -10
                '''
            }
        }

        stage('Generate Report') {
            steps {
                script {
                    echo "🚀 Starting automated report generation..."
                    
                    if (params.UPLOADED_EXCEL) {
                        echo "✅ Using uploaded file"
                        sh """
                            ./venv/bin/python generate_report.py "${params.UPLOADED_EXCEL}"
                        """
                    } else {
                        echo "📊 Auto-detecting Excel file from repository..."
                        sh """
                            ./venv/bin/python generate_report.py
                        """
                    }
                }
            }
        }

        stage('Send Email') {
            steps {
                script {
                    def htmlReport = readFile 'output/uptime_report.html'
                    
                    emailext(
                        subject: "Automated Uptime Report - ${new Date().format('dd-MMM-yyyy')}",
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

        stage('Archive') {
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
