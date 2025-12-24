pipeline {
    agent any

    stages {

        stage('Checkout Code') {
            steps {
                git branch: 'main',
                    url: 'https://github.com/vishalyadav74/uptime-report-pipeline.git'
            }
        }

        stage('Setup Python Environment') {
            steps {
                sh '''
                  python3 -m venv venv
                  ./venv/bin/pip install --upgrade pip
                  ./venv/bin/pip install -r requirements.txt
                '''
            }
        }

        stage('Detect Latest Excel File') {
            steps {
                sh '''
                  echo "🔍 Detecting latest Excel file from repo..."

                  LATEST_EXCEL=$(ls -t data/*.xlsx | head -n 1)

                  if [ -z "$LATEST_EXCEL" ]; then
                    echo "❌ No Excel file found in data/ folder"
                    exit 1
                  fi

                  echo "✅ Latest Excel file: $LATEST_EXCEL"
                  echo "LATEST_EXCEL=$LATEST_EXCEL" > excel.env
                '''
            }
        }

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  source excel.env
                  echo "📄 Using Excel file: $LATEST_EXCEL"

                  ./venv/bin/python generate_report.py "$LATEST_EXCEL"
                '''
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
            echo '✅ Uptime report generated successfully'
        }
        failure {
            echo '❌ Pipeline failed'
        }
    }
}
