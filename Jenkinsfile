pipeline {
    agent any

    parameters {
        file(
            name: 'UPTIME_EXCEL',
            description: 'Upload Uptime Excel File (Weekly + Quarterly sheets)'
        )
    }

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

        stage('Generate Uptime Report') {
            steps {
                sh '''
                  echo "Workspace: $WORKSPACE"
                  echo "Uploaded Excel file name: $UPTIME_EXCEL"

                  FILE_PATH="$WORKSPACE/$UPTIME_EXCEL"

                  echo "Resolved Excel file path: $FILE_PATH"
                  ls -l "$FILE_PATH"

                  export UPTIME_EXCEL="$FILE_PATH"
                  ./venv/bin/python generate_report.py
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
