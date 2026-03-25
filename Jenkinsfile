// 使用 Pipeline script from SCM 时，Git 凭据在 Jenkins 任务配置的 SCM 中设置（credentialsId: gitee-extapp）
// Docker 推送使用 docker-registry-credentials，若凭据 ID 不同请修改下方 credentialsId
// parameters 定义的 IMAGE_TAG 会在首次构建后自动生效，建议在任务中勾选「参数化构建过程」并添加 IMAGE_TAG 字符串参数以保持一致
pipeline {
    agent any

    environment {
        // 本地与仓库中的镜像短名（与 Dockerfile 所在目录为构建上下文）
        IMAGE_NAME = 'ohc-account-invoice'
        REGISTRY_IMAGE = 'release.daocloud.io/aigc/ohc-account-invoice'
    }

    parameters {
        string(name: 'IMAGE_TAG', defaultValue: 'latest', description: '镜像版本标签')
    }

    stages {
        stage('清理旧镜像') {
            steps {
                sh """
                    docker rmi -f ${env.IMAGE_NAME}:${params.IMAGE_TAG} || true
                    docker rmi -f ${env.REGISTRY_IMAGE}:${params.IMAGE_TAG} || true
                """
            }
        }

        stage('构建镜像') {
            steps {
                sh """
                    docker build --platform linux/amd64 -t ${env.IMAGE_NAME}:${params.IMAGE_TAG} .
                """
            }
        }

        stage('打标签') {
            steps {
                sh """
                    docker tag ${env.IMAGE_NAME}:${params.IMAGE_TAG} ${env.REGISTRY_IMAGE}:${params.IMAGE_TAG}
                """
            }
        }

        stage('推送镜像') {
            steps {
                withCredentials([usernamePassword(
                    credentialsId: 'docker-registry-credentials',
                    usernameVariable: 'DOCKER_USER',
                    passwordVariable: 'DOCKER_PASS'
                )]) {
                    sh """
                        echo \$DOCKER_PASS | docker login release.daocloud.io -u \$DOCKER_USER --password-stdin
                        docker push ${env.REGISTRY_IMAGE}:${params.IMAGE_TAG}
                    """
                }
            }
        }
    }
}
