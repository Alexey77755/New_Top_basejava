package ru.javawebinar.basejava.storage;

import ru.javawebinar.basejava.exception.StorageException;
import ru.javawebinar.basejava.model.Resume;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public abstract class AbstractFileStorage extends AbstractStorage<File> {
   private File directory ;
    protected AbstractFileStorage(File directory) {
     Objects.requireNonNull(directory,"directory must not be null");
     if(!directory.isDirectory()){
      throw new IllegalArgumentException(directory.getAbsolutePath() + "is not directory");
     }
     if(!directory.canRead() || !directory.canWrite()){
      throw new IllegalArgumentException(directory.getAbsolutePath() + "is not readable/writable");
     }
     this.directory=directory;
    }

    @Override
    protected File getSearchKey(String uuid) {
        return new File(directory, uuid);
    }

    @Override
    protected void doSave(Resume r, File file) {
     try {
      file.createNewFile();
      doWrite(r,file);
     } catch (IOException e) {
      throw new StorageException("IO error",file.getName(),e);
     }

    }

 protected abstract void doWrite(Resume r, File file) throws IOException;

 @Override
    protected void doUpdate(Resume r, File file) {
     try {
         doWrite(r,file);
     } catch (IOException e) {
         throw new StorageException("IO error",file.getName(),e);
     }
    }

    @Override
    protected boolean isExist(File file) {
        return file.exists();
    }

    @Override
    protected void doDelete(File file) {
        file.delete();
    }

    @Override
    protected Resume doGet(File file) {

        try {
            return doRead(file);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    protected abstract Resume doRead(File file) throws IOException;

    @Override
    protected List<Resume> doCopyAll()  {
        String[] list = directory.list();
        List<Resume> rList=new ArrayList<Resume>();
        if (list != null) {
            for (String name : list) {
                if (!new File( directory + "\\" + name).isDirectory()) {
                    try {
                        rList.add(doRead(new File( directory + "\\" + name)));
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                }
            }
        }

        return rList;
    }

    @Override
    public void clear() {
      //  File dir = new File(filePath);
        String[] list = directory.list();
        if (list != null) {
            for (String name : list) {
                if (!new File( directory + "\\" + name).isDirectory()) {
                    directory.delete();
                }
            }
        }
    }

    @Override
    public int size() {
        int counter = 0;

        String[] list = directory.list();
        if (list != null) {
            for (String name : list) {
                if (!new File(name).isDirectory()) {
                    counter++;
                }
            }

        }
        return counter;
    }
}
