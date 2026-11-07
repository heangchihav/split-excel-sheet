# Multi-stage Dockerfile for Spring Boot + Apache POI Excel Service
# Stage 1: Build application using Maven and Temurin JDK 21
FROM maven:3.9.9-eclipse-temurin-21 AS builder
WORKDIR /build

# Cache Maven dependencies
COPY pom.xml .
RUN mvn dependency:go-offline -B

# Build application JAR
COPY src ./src
RUN mvn clean package -DskipTests

# Stage 2: Production JRE runtime with Khmer fonts
FROM eclipse-temurin:21-jre-jammy AS runner
WORKDIR /app

# Install Khmer Unicode typography fonts and curl for healthchecks
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-khmeros \
    fonts-noto-extra \
    fonts-noto-ui-core \
    fontconfig \
    ca-certificates \
    curl \
    && fc-cache -f -v \
    && rm -rf /var/lib/apt/lists/*

# Run as non-root user for security
RUN groupadd -r spring && useradd -r -g spring spring
USER spring:spring

# Copy built JAR from builder stage
COPY --from=builder /build/target/split-excel-backend-*.jar app.jar

EXPOSE 8080

HEALTHCHECK --interval=15s --timeout=5s --start-period=30s --retries=3 \
  CMD curl -f http://localhost:8080/api/health || exit 1

ENTRYPOINT ["java", "-XX:+UseG1GC", "-XX:MaxRAMPercentage=75.0", "-jar", "app.jar"]

