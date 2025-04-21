package com.prollpa.serviceImpl;

import java.util.ArrayList;
import java.util.List;

import org.modelmapper.ModelMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.prollpa.entity.Role;
import com.prollpa.entity.User;
import com.prollpa.entity.UserRole;
import com.prollpa.entity.UserVscRole;
import com.prollpa.exception.ResourceNotFoundException;
import com.prollpa.payload.UserRoleDto;
import com.prollpa.repository.RoleRepository;
import com.prollpa.repository.UserRepository;
import com.prollpa.repository.UserRoleRepository;
import com.prollpa.service.UserRoleService;

@Service
public class UserRoleServiceImpl implements UserRoleService {
	private final UserRoleRepository userRoleRepository;
	private final ModelMapper modelMapper;
	private  final RoleRepository roleRepository;
	private final UserRepository userRepository;
	
	public UserRoleServiceImpl(UserRoleRepository userRoleRepository,ModelMapper modelMapper,RoleRepository roleRepository,UserRepository userRepository) {
		this.userRoleRepository=userRoleRepository;
		this.modelMapper=modelMapper;
		this.userRepository=userRepository;
		this.roleRepository=roleRepository;
	}

	@Override
	public UserRoleDto saveUserRole(UserRoleDto userDto) {
		UserRole userRole=new UserRole();
		Role role=roleRepository.findById(userDto.getRoleId()).orElseThrow(
				()-> new ResourceNotFoundException("roleNot found with this id "+userDto.getRoleId()));
		User user =userRepository.findById(userDto.getUserId()).orElseThrow(
				()-> new ResourceNotFoundException("user not found with this id "+userDto.getUserId()));
		userRole.setRole(role);
		userRole.setUser(user);
        UserRole saveUserRole = userRoleRepository.save(userRole);
        
		return mapUserRoletoUserRoleDto(saveUserRole);
	}

	@Override
	public List<UserRoleDto> getUserRoleByUserId(Long UserId) {
		List<UserRole> userRoleList = userRoleRepository.findByUser_userId(UserId);
		return mapUserRoletoUserRoleDtoList(userRoleList);
		
	}
	private UserRoleDto mapUserRoletoUserRoleDto(UserRole userRole){
		UserRoleDto userRoleDto=new UserRoleDto();
		userRoleDto.setUserRoleId(userRole.getUserRoleId());
		userRoleDto.setRoleId(userRole.getRole().getRoleId());
		userRoleDto.setUserId(userRole.getUser().getUserId());
		return userRoleDto;
		
	}
	private List<UserRoleDto> mapUserRoletoUserRoleDtoList(List<UserRole>userRolelist){
		
		List<UserRoleDto> userRoleDtoList=new ArrayList();
		for(UserRole userRole:userRolelist) {
			
			UserRoleDto userRoleDto=new UserRoleDto();
			userRoleDto.setUserRoleId(userRole.getUserRoleId());
			userRoleDto.setRoleId(userRole.getRole().getRoleId());
			userRoleDto.setUserId(userRole.getUser().getUserId());
			userRoleDtoList.add(userRoleDto);
		}
		return userRoleDtoList;
	}

}
