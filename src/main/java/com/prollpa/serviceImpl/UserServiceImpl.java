package com.prollpa.serviceImpl;

import java.util.Optional;


import org.modelmapper.ModelMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import com.prollpa.entity.User;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.UserDto;
import com.prollpa.repository.UserRepository;
import com.prollpa.service.UserService;

@Service
public class UserServiceImpl  implements UserService{
	@Autowired
	private UserRepository userRepository;
	@Autowired
	ModelMapper modelmapper;

	@Override
	public UserDto saveUser(UserDto dto) {
		User user=new User();
		user.setEmail(dto.getEmail());
		user.setName(dto.getName());
		user.setUsername(dto.getUsername());
		User saveUser = userRepository.save(user);
		return modelmapper.map(saveUser, UserDto.class);
		
	}

	@Override
	public UserDto getUserByUsername(String username) {
		User user = userRepository.findByUsername(username).get();
		return modelmapper.map(user,UserDto.class);
				
	}

	@Override
	public UserDto getUserByUserId(long userId) {
		User user = userRepository.findById(userId).orElseThrow(
				()-> new ResourceNotFoundException("user not found with this id "+userId));
		return UserDto.builder().userId(user.getUserId()).name(user.getName()).email(user.getEmail()).username(user.getUsername()).build();
	}
	

}
