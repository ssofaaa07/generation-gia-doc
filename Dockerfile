# Используем официальный образ
FROM eclipse-temurin:20-jdk-jammy

  # Рабочая директория
WORKDIR /app
  
  # Копируем Gradle-файлы для кэширования
COPY gradlew .
COPY gradle gradle
COPY build.gradle .
COPY settings.gradle .
COPY src src

  # Запускаем сборку
RUN ./gradlew build --no-daemon
  
  # Команда для запуска
CMD ["java", "-jar", "build/libs/generation-gia-doc-0.0.1-SNAPSHOT.jar"]