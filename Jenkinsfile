pipeline {
    agent any

    parameters {
        file(name: 'EXCEL_FILE', description: 'Select Excel file for uptime report')
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
                    // If a file parameter was provided, copy it to workspace
                    if (params.EXCEL_FILE) {
                        echo "📥 Using uploaded file: ${params.EXCEL_FILE}"
                        sh '''
                          mkdir -p input
                          cp "$EXCEL_FILE" input/uptime.xlsx
                        '''
                        env.UPTIME_EXCEL = "${WORKSPACE}/input/uptime.xlsx"
                    } else {
                        // Fallback to the file in repository
                        echo "📄 Using default file from repository"
                        env.UPTIME_EXCEL = "${WORKSPACE}/uptime_latest1.xlsx"
                    }
                    
                    echo "✅ Excel file set to: ${UPTIME_EXCEL}"
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
                    def htmlReport = readFile 'output/uptime_report.html'

                    emailext(
                        subject: "SaaS Application Uptime Report – Weekly & Quarterly",
                        body: htmlReport,
                        mimeType: 'text/html',
                        to: env.EMAIL_TO,
                        from: 'Jenkins <yv741518@gmail.com>',
                        replyTo: 'yv741518@gmail.com'
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
            echo '✅ Report generated and email sent successfully'
        }
        failure {
            echo '❌ Pipeline failed'
        }
    }
}
