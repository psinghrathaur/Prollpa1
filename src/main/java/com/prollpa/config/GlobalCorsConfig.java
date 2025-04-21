package com.prollpa.config;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.CorsRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

@Configuration
public class GlobalCorsConfig implements WebMvcConfigurer {

    @Override
    public void addCorsMappings(CorsRegistry registry) {
        // This applies to all endpoints ("/**")
        registry.addMapping("/**")                          // Apply to all paths
                .allowedOrigins("http://localhost:4200","http://localhost:8085","http://localhost:8090","http://localhost/127.0.0.1:8085","https://saddlebrown-whale-122097.hostingersite.com")    // Allow requests from this frontend URL
                .allowedMethods("GET", "POST", "PUT", "DELETE", "OPTIONS")  // Allow these methods
                .allowedHeaders("Content-Type", "Authorization")  // Allow these headers in the request
                .allowCredentials(true); // Allow sending credentials (cookies, HTTP authentication)
    }
    
}
