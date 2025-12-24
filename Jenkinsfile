pipeline {
    agent any

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

        stage('Detect Latest Excel File') {
            steps {
                sh '''
                  echo "🔍 Searching latest Excel file..."
                  LATEST_EXCEL=$(ls -t data/*.xlsx | head -n 1)

                  if [ -z "$LATEST_EXCEL" ]; then
                    echo "❌ No Excel file found in data folder"
                    exit 1
                  fi

                  echo "✅ Using Excel file: $LATEST_EXCEL"
                  echo "EXCEL_FILE=$LATEST_EXCEL" >> $GITHUB_ENV
                '''
            }
        }

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  echo "📊 Generating report from $EXCEL_FILE"
                  ./venv/bin/python generate_report.py "$EXCEL_FILE"
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
