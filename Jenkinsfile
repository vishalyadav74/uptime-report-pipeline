pipeline {
    agent any

    parameters {
        string(name: 'MAIL_TO', defaultValue: '', description: 'To (comma separated)')
        string(name: 'MAIL_CC', defaultValue: '', description: 'CC (comma separated)')
    }

    stages {

        stage('Setup') {
            steps {
                sh '''
                  python3 -m venv venv
                  ./venv/bin/pip install -r requirements.txt
                '''
            }
        }

        stage('Generate Report') {
            steps {
                sh "./venv/bin/python generate_report.py"
            }
        }

        stage('Send Email') {
            steps {
                withCredentials([
                  usernamePassword(
                    credentialsId: 'ITSM_SMTP',
                    usernameVariable: 'SMTP_USER',
                    passwordVariable: 'SMTP_PASSWORD'
                  )
                ]) {
                    sh """
                      ./venv/bin/python send.py \
                        --subject "SAAS Accounts Weekly & Quarterly Application Uptime Report" \
                        --to "${params.MAIL_TO}" \
                        --cc "${params.MAIL_CC}" \
                        --body output/uptime_report.html
                    """
                }
            }
        }

        stage('Archive') {
            steps {
                archiveArtifacts artifacts: 'output/uptime_report.html'
            }
        }
    }
}
