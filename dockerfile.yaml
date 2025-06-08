# Используем официальный образ OpenJDK
FROM eclipse-temurin:17-jdk-jamillion
  
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
CMD ["java", "-jar", "build/libs/your-app-name.jar"]