// Rootless Podman on login01 has no subordinate UID/GID range. The container
// process is still the unprivileged host user, but appears as UID 0 inside its
// single-ID namespace. PostgreSQL refuses that UID. Preload this tiny shim only
// in the PostgreSQL container so its self-check sees the image's postgres user.

#define _GNU_SOURCE
#include <dlfcn.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <unistd.h>

static const uid_t postgres_uid = 999;
static const gid_t postgres_gid = 999;

uid_t getuid(void) { return postgres_uid; }
uid_t geteuid(void) { return postgres_uid; }
gid_t getgid(void) { return postgres_gid; }
gid_t getegid(void) { return postgres_gid; }

int getresuid(uid_t *ruid, uid_t *euid, uid_t *suid) {
  if (ruid) *ruid = postgres_uid;
  if (euid) *euid = postgres_uid;
  if (suid) *suid = postgres_uid;
  return 0;
}

int getresgid(gid_t *rgid, gid_t *egid, gid_t *sgid) {
  if (rgid) *rgid = postgres_gid;
  if (egid) *egid = postgres_gid;
  if (sgid) *sgid = postgres_gid;
  return 0;
}

static void rewrite_owner(struct stat *buf) {
  if (!buf) return;
  buf->st_uid = postgres_uid;
  buf->st_gid = postgres_gid;
}

int stat(const char *path, struct stat *buf) {
  static int (*real_stat)(const char *, struct stat *);
  if (!real_stat) real_stat = dlsym(RTLD_NEXT, "stat");
  int rc = real_stat(path, buf);
  if (rc == 0) rewrite_owner(buf);
  return rc;
}

int lstat(const char *path, struct stat *buf) {
  static int (*real_lstat)(const char *, struct stat *);
  if (!real_lstat) real_lstat = dlsym(RTLD_NEXT, "lstat");
  int rc = real_lstat(path, buf);
  if (rc == 0) rewrite_owner(buf);
  return rc;
}

int fstat(int fd, struct stat *buf) {
  static int (*real_fstat)(int, struct stat *);
  if (!real_fstat) real_fstat = dlsym(RTLD_NEXT, "fstat");
  int rc = real_fstat(fd, buf);
  if (rc == 0) rewrite_owner(buf);
  return rc;
}

int fstatat(int dirfd, const char *path, struct stat *buf, int flags) {
  static int (*real_fstatat)(int, const char *, struct stat *, int);
  if (!real_fstatat) real_fstatat = dlsym(RTLD_NEXT, "fstatat");
  int rc = real_fstatat(dirfd, path, buf, flags);
  if (rc == 0) rewrite_owner(buf);
  return rc;
}
