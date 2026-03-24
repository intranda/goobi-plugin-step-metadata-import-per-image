
pipeline {

  agent {
    docker {
      image 'maven:3-eclipse-temurin-21'
      args '-v $HOME/.m2:/var/maven/.m2:z -v $HOME/.config:/var/maven/.config -v $HOME/.sonar:/var/maven/.sonar -u 1000 -ti -e _JAVA_OPTIONS=-Duser.home=/var/maven -e MAVEN_CONFIG=/var/maven/.m2'
    }
  }

  options {
    buildDiscarder logRotator(artifactDaysToKeepStr: '', artifactNumToKeepStr: '15', daysToKeepStr: '90', numToKeepStr: '')
  }

  stages {
    stage('prepare') {
      steps {
        sh 'git reset --hard HEAD && git clean -fdx'
      }
    }
    stage('build-snapshot') {
      when {
        not {
          anyOf {
            branch 'master'
            branch 'release_*'
            branch 'hotfix_release_*'
            branch 'sonar_*'
            allOf {
              branch 'PR-*'
              expression { env.CHANGE_BRANCH.startsWith("release_") }
            }
          }
        }
      }
      steps {
        sh 'mvn clean verify -U -P snapshot-build'
      }
    }
    stage('build-release') {
      when {
        anyOf {
          branch 'master'
          branch 'release_*'
          branch 'hotfix_release_*'
          allOf {
            branch 'PR-*'
            expression { env.CHANGE_BRANCH.startsWith("release_") }
          }
        }
      }
      parallel {

        stage('test') {
          agent {
            docker {
              image mavenDockerImage
              args mavenDockerArgs
              reuseNode true
            }
          }
          steps {
            script {
              def strict = env.BRANCH_NAME == 'master'
              def cmd = "mvn test -Dmaven.main.skip=true -Drevision=\$BUILD_VERSION -P '!local-development' --no-transfer-progress"
              if (strict) {
                sh cmd
              } else {
                catchError(buildResult: 'UNSTABLE', stageResult: 'UNSTABLE') {
                  sh cmd
                }
              }
            }
            junit allowEmptyResults: true, testResults: '**/target/surefire-reports/*.xml'
            step([
                    $class           : 'JacocoPublisher',
                    execPattern      : '**/target/jacoco.exec',
                    classPattern     : '**/target/classes/',
                    sourcePattern    : '**/src/main/java',
                    exclusionPattern : '**/*Test.class'
            ])
          }
        }

        stage('checkstyle') {
          agent {
            docker {
              image mavenDockerImage
              args mavenDockerArgs
              reuseNode true
            }
          }
          steps {
            script {
              def strict = (env.BRANCH_NAME == 'master') && !env.NO_STRICT_CHECKSTYLE
              def cmd = "mvn checkstyle:check -Drevision=\$BUILD_VERSION -P '!local-development' --no-transfer-progress"
              if (strict) {
                sh cmd
              } else {
                catchError(buildResult: 'UNSTABLE', stageResult: 'UNSTABLE') {
                  sh cmd
                }
              }
            }
            recordIssues(
                    id: 'checkstyle-plugin',
                    tools: [checkStyle(pattern: '**/target/checkstyle-result.xml')],
                    qualityGates: [[threshold: 1, type: 'TOTAL', unstable: true]]
            )
          }
        }

      }
    }
    stage('build-sonar') {
      when {
        branch 'sonar_*'
      }
      steps {
        sh 'mvn clean verify -U -P sonar-build'
      }
    }
    stage('sonarcloud') {
      when {
        allOf {
          anyOf {
            branch 'master'
            branch 'release_*'
            branch 'hotfix_release_*'
            branch 'sonar_*'
            allOf {
              branch 'PR-*'
              expression { env.CHANGE_BRANCH.startsWith("release_") }
            }
          }
          not {
            expression {
              return fileExists('DO_NOT_PUBLISH')
            }
          }
        }
      }
      steps {
        withCredentials([string(credentialsId: 'jenkins-sonarcloud', variable: 'TOKEN')]) {
          sh 'mvn verify sonar:sonar -Dsonar.token=$TOKEN -U'
        }
      }
    }
    stage('deploy-libs') {
      when {
        anyOf {
          branch 'master'
          branch 'develop'
          branch 'hotfix_release_*'
        }
      }
      steps {
        script {
          if (fileExists('module-lib/pom.xml')) {
            def altRepo = fileExists('DO_NOT_PUBLISH')
                    ? "-DaltDeploymentRepository=\$NEXUS_INTERNAL_REPO -DaltSnapshotDeploymentRepository=\$NEXUS_INTERNAL_REPO"
                    : ''
            sh "mvn -N deploy -Dmaven.main.skip=true -Dmaven.test.skip=true -Drevision=\$BUILD_VERSION -U ${altRepo} --no-transfer-progress"
            sh "mvn -f module-lib/pom.xml deploy -Dmaven.main.skip=true -Dmaven.test.skip=true -Drevision=\$BUILD_VERSION -U ${altRepo} --no-transfer-progress"
          }
        }
      }
    }
    stage('tag release') {
      when {
        anyOf {
          branch 'master'
          branch 'hotfix_release_*'
        }
      }
      agent {
        docker {
          image mavenDockerImage
          args mavenDockerArgs
          reuseNode true
        }
      }
      steps {
        withCredentials([gitUsernamePassword(credentialsId: '93f7e7d3-8f74-4744-a785-518fc4d55314',
                 gitToolName: 'git-tool')]) {
          sh '''#!/bin/bash -xe
            PLUGIN_NAME=$(basename $(git remote get-url origin) .git)

            if [ -f "DO_NOT_PUBLISH" ]; then
              REPO_URL="$COLLECTION_REPO_URL"
              SUBMODULE_PATH="private-plugins/$PLUGIN_NAME"
              BRANCH="master"
            else
              REPO_URL="$CORE_REPO_URL"
              SUBMODULE_PATH="plugins/$PLUGIN_NAME"
              BRANCH="develop"
            fi

            WORK_DIR=$(mktemp -d)
            git clone --depth 1 --branch $BRANCH "$REPO_URL" "$WORK_DIR"
            cd "$WORK_DIR"
            git submodule update --init --remote -- "$SUBMODULE_PATH"
            if git status --porcelain -- "$SUBMODULE_PATH" | grep -q .; then
              git add "$SUBMODULE_PATH"
              git commit -m "Update ${PLUGIN_NAME} to latest master"
              git push origin $BRANCH
            else
              echo "Submodule already up to date."
            fi
            rm -rf "$WORK_DIR"
          '''
        }
      }
    }
  }

  post {
    always {
      junit allowEmptyResults: true, testResults: "**/target/surefire-reports/*.xml"
      step([
        $class           : 'JacocoPublisher',
        execPattern      : '**/target/jacoco.exec',
        classPattern     : '**/target/classes/',
        sourcePattern    : '**/src/main/java',
        exclusionPattern : '**/*Test.class'
      ])
      recordIssues (
        enabledForFailure: true, aggregatingResults: false,
        tools: [checkStyle(pattern: 'target/checkstyle-result.xml', reportEncoding: 'UTF-8')]
      )
    }
    success {
      archiveArtifacts artifacts: '**/target/*.jar, install/*', fingerprint: true, onlyIfSuccessful: true
    }
    changed {
      emailext(
              subject: '${DEFAULT_SUBJECT}',
              body: '${DEFAULT_CONTENT}',
              recipientProviders: [requestor(), culprits()],
              attachLog: true
      )
    }
  }
}
