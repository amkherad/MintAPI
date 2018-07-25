/* -*- Mode: C; tab-width: 4 -*- */

/*
 *	Author: Sam Rushing <rushing@nightmare.com>
 *	$Id: scm_socket.h,v 1.1 1997/03/10 07:42:11 rushing Exp $
 */

/* minimal socket interface for gambit-c scheme */

#ifdef _WIN32
#  include <winsock.h>
#  define EINPROGRESS	WSAEINPROGRESS
#  define EWOULDBLOCK	WSAEWOULDBLOCK
#  define EALREADY		WSAEALREADY
#  define ECONNRESET	WSAECONNRESET
#  define alloca		_alloca
#  define close			closesocket
#else
#  include <unistd.h>
#  include <sys/time.h>
#  include <sys/types.h>
#  include <sys/socket.h>
#  include <netdb.h>
#  include <netinet/in.h>
#  include <sys/un.h>
#  include <errno.h>
#endif

#include <stdio.h>

typedef struct {
  int fd;
  int family;
  int type;
  int proto;
  union {
    struct sockaddr_in in;
#ifndef _WIN32
    struct sockaddr_un un;
#endif
  } sock_addr;
} scm_socket;

int
in_address_as_string (struct sockaddr_in * addr, char * buffer)
{
  long ip = ntohl (addr->sin_addr.s_addr);
  int port = ntohs (addr->sin_port);
  return sprintf (
    buffer, "%d.%d.%d.%d:%d", 
	(int) (ip>>24) & 0xff, (int) (ip>>16) & 0xff,
	(int) (ip>> 8) & 0xff, (int) (ip>> 0) & 0xff,
	port
	);
}

int
in_address_from_string (struct sockaddr_in * addr, char * buffer)
{
  int a,b,c,d,p;
  if (sscanf (buffer, "%d.%d.%d.%d:%d", &a, &b, &c, &d, &p) == 5) {
	addr->sin_family = AF_INET;
	addr->sin_addr.s_addr = htonl (
          ((long) a << 24) | ((long) b << 16) |
          ((long) c << 8) | ((long) d << 0)
        );
	addr->sin_port = htons (p);
	return 1;
  } else {
	return 0;
  }
}


scm_socket *
new_scm_socket (int family, int type, int protocol)
{
  int fd;
  fd = socket (family, type, protocol);
  if (fd) {
	scm_socket * s = (scm_socket *) malloc (sizeof (scm_socket));
	s->fd = fd;
	s->family = family;
	s->type = type;
	s->proto = protocol;
	return s;
  } else {
	return NULL;
  }
}

int
scm_socket_bind (scm_socket * s, char * addr_string)
{
  if (s->family == AF_INET) {
	if (!in_address_from_string ((struct sockaddr_in *) &((s->sock_addr).in), addr_string)) {
	  return 0;
	}
  } else {
	/* for now, forget about AF_UNIX */
	return 0;
  }
  return bind (s->fd, (struct sockaddr *) &(s->sock_addr), sizeof(struct sockaddr_in));
}

scm_socket *
scm_socket_accept (scm_socket * s)
{
  int result;
  int addr_len = sizeof (struct sockaddr_in);
  scm_socket * client = (scm_socket *) malloc (sizeof (scm_socket));

  result = accept (
      s->fd,
	  (struct sockaddr *)&(client->sock_addr),
	  &addr_len
	  );

  if (result != -1) {
	client->fd		= result;
	client->family	= s->family;
	client->type	= s->type;
	client->proto	= s->proto;
	return client;
  } else {
	free (client);
	return NULL;
  }
}

int
scm_socket_connect (scm_socket * s, char * addr_string)
{
  int result;
  if (s->family == AF_INET) {
	if (!in_address_from_string ((struct sockaddr_in *) &((s->sock_addr).in), addr_string)) {
	  return -2;
	}
  } else {
	/* for now, forget about AF_UNIX */
	fprintf (stderr, "address family %d not supported\n", s->type);
	return -3;
  }
  /* do the connect */
  if (connect (s->fd, (struct sockaddr *) &(s->sock_addr), sizeof(struct sockaddr_in))) {
	return errno;
  } else {
	return 0;
  }
}

#ifndef _WIN32
#  include <sys/fcntl.h>
#endif

void
scm_socket_set_blocking (scm_socket *s, int block)
{
#ifndef _WIN32
  int delay_flag;
  delay_flag = fcntl (s->fd, F_GETFL, 0);
  if (block) {
	delay_flag &= (~O_NDELAY);
  } else {
	delay_flag |= O_NDELAY;
  }
  fcntl (s->fd, F_SETFL, delay_flag);
#else
  block = !block;
  ioctlsocket(s->fd, FIONBIO, (u_long*)&block);
#endif
}

/* On Win32, we need to initialize the WINSOCK library before we can use it. */

int
scm_socket_initialize (void)
{
#ifdef _WIN32
  WSADATA wsa_data;
  return (WSAStartup (0x0101, &wsa_data) == 0);
#else
  return 1;
#endif
}
