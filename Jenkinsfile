pipeline {
    agent any

    parameters {
        file(name: 'UPTIME_EXCEL', description: 'Upload Uptime Excel File')
    }

    stages {

        stage('Checkout Code') {
            steps {
                git branch: 'main',
                    url: 'https://github.com/vishalyadav74/uptime-report-pipeline.git'
            }
        }

        stage('Setup Python') {
            steps {
                sh '''
                  python3 -m venv venv
                  ./venv/bin/pip install --upgrade pip
                  ./venv/bin/pip install -r requirements.txt
                '''
            }
        }

        stage('Generate Report') {
            steps {
                sh '''
                  echo "Uploaded file: ${params.UPTIME_EXCEL}"
                  export UPTIME_EXCEL=${WORKSPACE}/${params.UPTIME_EXCEL}
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
